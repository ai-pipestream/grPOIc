package ai.pipestream.grpoic.server;

import ai.pipestream.grpoic.parse.DocumentParser;
import ai.pipestream.grpoic.parse.DocumentTooLargeException;
import ai.pipestream.grpoic.parse.InvalidDocumentException;
import ai.pipestream.grpoic.parse.PoiLimits;
import ai.pipestream.grpoic.parse.ProtectedDocumentException;
import ai.pipestream.grpoic.parse.UnsupportedFormatException;
import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.GetServiceInfoRequest;
import ai.pipestream.poi.v1.GetServiceInfoResponse;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import ai.pipestream.poi.v1.UiInfo;
import com.google.protobuf.ByteString;
import io.grpc.Status;
import io.grpc.stub.ServerCallStreamObserver;
import io.grpc.stub.StreamObserver;
import java.time.Duration;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Condition;
import java.util.concurrent.locks.ReentrantLock;
import java.util.function.Consumer;
import org.apache.poi.Version;

/**
 * The gRPC face over DocumentParser. Each call runs start to finish on its
 * own virtual thread: admission, upload, parse, events.
 *
 * <p><b>Admission before buffering.</b> A call takes a parse slot before it
 * accepts its first chunk, and pulls chunks one at a time (inbound flow
 * control), so the semaphore bounds the documents held in memory, not just
 * the parses running. A queued client's bytes stay in its transport window.
 * Chunks are kept as the received ByteStrings, joined without copying, and
 * the parser reads that buffer in place. A client that stops sending for
 * {@link #UPLOAD_IDLE_TIMEOUT} gives its slot back.
 *
 * <p><b>Backpressure out.</b> Events are written only while the transport
 * is ready for more; otherwise the parse thread waits for the ready signal.
 * A slow client stalls its parse instead of queueing unbounded messages.
 *
 * <p><b>Cancellation.</b> A cancelled or expired call stops waiting for a
 * slot, stops its upload, and stops its parse at the next event; it counts
 * as neither parsed nor failed.
 */
public final class PoiParseServiceImpl extends PoiParseServiceGrpc.PoiParseServiceImplBase {

  public static final String SERVICE_VERSION = "0.1.0";
  public static final String API_VERSION = "v1";
  // Advertised to the shared demo shell so it can mount this service's tab.
  public static final UiInfo UI_INFO = UiInfo.newBuilder()
      .setTitle("POI")
      .setPath("/ui/grpoic")
      .setDescription("Apache POI wrapper for office documents")
      .build();

  /** How long an admitted upload may go without a chunk before it is dropped. */
  static final Duration UPLOAD_IDLE_TIMEOUT = Duration.ofSeconds(30);

  // Built ahead of need: when the heap is exhausted, failing the call
  // should allocate as little as possible.
  private static final Status OUT_OF_MEMORY = Status.RESOURCE_EXHAUSTED
      .withDescription("document needs more memory than the server can give it");
  private static final Status TOO_DEEP = Status.RESOURCE_EXHAUSTED
      .withDescription("document structure nests too deeply to parse");

  /** The parse step, separate so tests can fail it in ways no document reliably does. */
  @FunctionalInterface
  interface Parser {
    void parse(String documentId, ByteString data, Consumer<ParseEvent> emit);
  }

  private final long maxDocumentBytes;
  private final int maxConcurrentParses;
  private final Semaphore parseSlots;
  private final ExecutorService executor;
  private final Parser parser;
  private final Duration uploadIdleTimeout;
  private final ParseCounters counters = new ParseCounters();

  public PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses,
                             ExecutorService executor) {
    this(maxDocumentBytes, maxConcurrentParses, executor, DocumentParser::parse,
        UPLOAD_IDLE_TIMEOUT);
  }

  PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses, ExecutorService executor,
                      Parser parser) {
    this(maxDocumentBytes, maxConcurrentParses, executor, parser, UPLOAD_IDLE_TIMEOUT);
  }

  PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses, ExecutorService executor,
                      Parser parser, Duration uploadIdleTimeout) {
    this.maxDocumentBytes = maxDocumentBytes;
    this.maxConcurrentParses = maxConcurrentParses;
    this.parseSlots = new Semaphore(maxConcurrentParses);
    this.executor = executor;
    this.parser = parser;
    this.uploadIdleTimeout = uploadIdleTimeout;
    // POI's limits are JVM-wide; they follow the cap of the server in it.
    PoiLimits.install(maxDocumentBytes);
  }

  public ParseCounters counters() {
    return counters;
  }

  @Override
  public StreamObserver<ParseRequestChunk> parseDocument(StreamObserver<ParseEvent> observer) {
    // grpc-java hands every server call a ServerCallStreamObserver; flow
    // control and cancellation are only reachable through it, and only
    // before this method returns.
    ServerCallStreamObserver<ParseEvent> responses =
        (ServerCallStreamObserver<ParseEvent>) observer;
    ParseCall call = new ParseCall(responses);
    responses.disableAutoRequest();
    responses.setOnReadyHandler(call::wake);
    responses.setOnCancelHandler(call::cancel);
    executor.execute(call::run);
    return call;
  }

  /** What the transport delivered; the call's thread takes these in order. */
  private sealed interface Inbound {}

  private record Chunk(ParseRequestChunk message) implements Inbound {}

  private enum Signal implements Inbound { HALF_CLOSED, ABORTED }

  /** The client went away; the parse unwinds without answering. */
  private static final class CallCancelled extends RuntimeException {
    CallCancelled() {
      super("call cancelled", null, false, false);
    }
  }

  /**
   * One ParseDocument call. The StreamObserver methods run on transport
   * threads and only enqueue; {@link #run()} does the work on the call's
   * virtual thread.
   */
  private final class ParseCall implements StreamObserver<ParseRequestChunk> {
    private final ServerCallStreamObserver<ParseEvent> responses;
    private final BlockingQueue<Inbound> inbound = new LinkedBlockingQueue<>();
    private final ReentrantLock lock = new ReentrantLock();
    private final Condition writable = lock.newCondition();
    private boolean signalled;
    private volatile boolean cancelled;
    private String documentId = "";

    ParseCall(ServerCallStreamObserver<ParseEvent> responses) {
      this.responses = responses;
    }

    @Override
    public void onNext(ParseRequestChunk chunk) {
      inbound.add(new Chunk(chunk));
    }

    @Override
    public void onError(Throwable error) {
      inbound.add(Signal.ABORTED);
    }

    @Override
    public void onCompleted() {
      inbound.add(Signal.HALF_CLOSED);
    }

    void cancel() {
      cancelled = true;
      inbound.add(Signal.ABORTED);
      wake();
    }

    void wake() {
      lock.lock();
      try {
        signalled = true;
        writable.signalAll();
      } finally {
        lock.unlock();
      }
    }

    void run() {
      try {
        if (!admit()) return;
      } catch (InterruptedException interrupt) {
        Thread.currentThread().interrupt();
        fail(Status.UNAVAILABLE.withDescription("server shutting down"));
        return;
      }
      try {
        ByteString data = receive();
        if (data != null) parse(data);
      } catch (InterruptedException interrupt) {
        Thread.currentThread().interrupt();
        fail(Status.UNAVAILABLE.withDescription("server shutting down"));
      } catch (RuntimeException unexpected) {
        if (!cancelled) {
          counters.recordFailed();
          fail(Status.INTERNAL.withDescription("upload fault: " + unexpected.getMessage()));
        }
      } finally {
        parseSlots.release();
      }
    }

    /** Waits for a parse slot; false when the call ended first. */
    private boolean admit() throws InterruptedException {
      while (!parseSlots.tryAcquire(100, TimeUnit.MILLISECONDS)) {
        if (cancelled) return false;
      }
      if (cancelled) {
        parseSlots.release();
        return false;
      }
      return true;
    }

    /**
     * Pulls chunks one at a time under the byte cap. Null when the call was
     * rejected (the status is sent) or abandoned by the client.
     */
    private ByteString receive() throws InterruptedException {
      ByteString received = ByteString.EMPTY;
      boolean sawComplete = false;
      responses.request(1);
      while (true) {
        Inbound next = inbound.poll(uploadIdleTimeout.toMillis(), TimeUnit.MILLISECONDS);
        if (next == null) {
          counters.recordRejected();
          fail(Status.DEADLINE_EXCEEDED.withDescription(
              "no upload data for " + uploadIdleTimeout.toSeconds() + " s"));
          return null;
        }
        if (next instanceof Chunk(ParseRequestChunk chunk)) {
          if (documentId.isEmpty() && !chunk.getDocumentId().isEmpty()) {
            documentId = chunk.getDocumentId();
          }
          if (received.size() + (long) chunk.getData().size() > maxDocumentBytes) {
            counters.recordRejected();
            fail(Status.RESOURCE_EXHAUSTED
                .withDescription("document exceeds " + maxDocumentBytes + " bytes"));
            return null;
          }
          received = received.concat(chunk.getData());
          if (chunk.getComplete()) sawComplete = true;
          responses.request(1);
        } else if (next == Signal.HALF_CLOSED) {
          if (!sawComplete || received.isEmpty()) {
            counters.recordRejected();
            fail(Status.INVALID_ARGUMENT.withDescription(received.isEmpty()
                ? "no document bytes received"
                : "stream ended without a chunk marked complete"));
            return null;
          }
          return received;
        } else {
          return null;
        }
      }
    }

    private void parse(ByteString data) {
      try {
        parser.parse(documentId, data, this::emit);
        complete();
        counters.recordParsed();
      } catch (CallCancelled abandoned) {
        // nothing to answer; the client is gone
      } catch (UnsupportedFormatException unsupported) {
        counters.recordRejected();
        fail(Status.UNIMPLEMENTED.withDescription(unsupported.getMessage()));
      } catch (InvalidDocumentException invalid) {
        counters.recordRejected();
        fail(Status.INVALID_ARGUMENT.withDescription(invalid.getMessage()));
      } catch (ProtectedDocumentException encrypted) {
        counters.recordRejected();
        fail(Status.FAILED_PRECONDITION.withDescription(encrypted.getMessage()));
      } catch (DocumentTooLargeException tooLarge) {
        counters.recordRejected();
        fail(Status.RESOURCE_EXHAUSTED.withDescription(tooLarge.getMessage()));
      } catch (Exception unexpected) {
        if (cancelled) return;
        counters.recordFailed();
        fail(Status.INTERNAL.withDescription("parser fault: " + unexpected.getMessage()));
      } catch (OutOfMemoryError exhausted) {
        // The document's allocations become unreachable as the stack unwinds,
        // so the call can still be failed; rethrown so the log shows it.
        counters.recordFailed();
        fail(OUT_OF_MEMORY);
        throw exhausted;
      } catch (StackOverflowError tooDeep) {
        counters.recordFailed();
        fail(TOO_DEEP);
      } catch (Throwable fault) {
        // Any other Error (a linkage failure, an assertion) must still close
        // the call: an open call leaves the client waiting for its deadline.
        counters.recordFailed();
        fail(Status.INTERNAL
            .withDescription("parser fault: " + fault.getClass().getSimpleName()));
        if (fault instanceof VirtualMachineError machineError) throw machineError;
      }
    }

    /** Writes one event once the transport can take it. */
    private void emit(ParseEvent event) {
      try {
        // isReady() is asked outside the lock the ready handler takes, so
        // the transport and this thread never wait on each other; the flag
        // keeps a signal that lands between the check and the wait.
        while (!cancelled && !responses.isReady()) {
          lock.lock();
          try {
            if (!signalled) writable.await(1, TimeUnit.SECONDS);
            signalled = false;
          } finally {
            lock.unlock();
          }
        }
      } catch (InterruptedException interrupt) {
        Thread.currentThread().interrupt();
        throw new CallCancelled();
      }
      if (cancelled) throw new CallCancelled();
      try {
        responses.onNext(event);
      } catch (RuntimeException closedUnderneath) {
        if (cancelled || responses.isCancelled()) throw new CallCancelled();
        throw closedUnderneath;
      }
    }

    private void complete() {
      try {
        responses.onCompleted();
      } catch (RuntimeException closedUnderneath) {
        if (!responses.isCancelled()) throw closedUnderneath;
        throw new CallCancelled();
      }
    }

    /** Fails the call; a call the client already abandoned has nothing left to fail. */
    private void fail(Status status) {
      try {
        responses.onError(status.asRuntimeException());
      } catch (RuntimeException alreadyClosed) {
        // cancelled or closed underneath us
      }
    }
  }

  @Override
  public void getServiceInfo(GetServiceInfoRequest request,
                             StreamObserver<GetServiceInfoResponse> responses) {
    responses.onNext(
        GetServiceInfoResponse.newBuilder()
            .setServiceVersion(SERVICE_VERSION)
            .setPoiVersion(Version.getVersion())
            .setApiVersion(API_VERSION)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_DOCX)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_XLSX)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_PPTX)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_LEGACY_DOC)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_LEGACY_XLS)
            .addSupportedFormats(DocumentFormat.DOCUMENT_FORMAT_LEGACY_PPT)
            .setMaxDocumentBytes(maxDocumentBytes)
            .setMaxConcurrentParses(maxConcurrentParses)
            .setUi(UI_INFO)
            .build());
    responses.onCompleted();
  }
}
