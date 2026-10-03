package ai.pipestream.grpoic.server;

import ai.pipestream.grpoic.parse.DocumentParser;
import ai.pipestream.grpoic.parse.InvalidDocumentException;
import ai.pipestream.grpoic.parse.UnsupportedFormatException;
import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.GetServiceInfoRequest;
import ai.pipestream.poi.v1.GetServiceInfoResponse;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import ai.pipestream.poi.v1.UiInfo;
import io.grpc.Status;
import io.grpc.stub.StreamObserver;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Semaphore;
import java.util.function.Consumer;
import org.apache.poi.Version;

/**
 * The gRPC face over DocumentParser. Chunks accumulate in memory under a hard
 * byte cap; a completed upload parses on a virtual thread, with a semaphore
 * bounding concurrent parses (POI documents are single-threaded and
 * memory-hungry; distinct documents on distinct threads is the supported
 * pattern).
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

  // Built ahead of need: when the heap is exhausted, failing the call
  // should allocate as little as possible.
  private static final Status OUT_OF_MEMORY = Status.RESOURCE_EXHAUSTED
      .withDescription("document needs more memory than the server can give it");
  private static final Status TOO_DEEP = Status.RESOURCE_EXHAUSTED
      .withDescription("document structure nests too deeply to parse");

  /** The parse step, separate so tests can fail it in ways no document reliably does. */
  @FunctionalInterface
  interface Parser {
    void parse(String documentId, byte[] bytes, Consumer<ParseEvent> emit);
  }

  private final long maxDocumentBytes;
  private final int maxConcurrentParses;
  private final Semaphore parseSlots;
  private final ExecutorService executor;
  private final Parser parser;
  private final ParseCounters counters = new ParseCounters();

  public PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses,
                             ExecutorService executor) {
    this(maxDocumentBytes, maxConcurrentParses, executor, DocumentParser::parse);
  }

  PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses, ExecutorService executor,
                      Parser parser) {
    this.maxDocumentBytes = maxDocumentBytes;
    this.maxConcurrentParses = maxConcurrentParses;
    this.parseSlots = new Semaphore(maxConcurrentParses);
    this.executor = executor;
    this.parser = parser;
  }

  public ParseCounters counters() {
    return counters;
  }

  @Override
  public StreamObserver<ParseRequestChunk> parseDocument(StreamObserver<ParseEvent> responses) {
    return new StreamObserver<>() {
      private final ByteArrayOutputStream buffer = new ByteArrayOutputStream();
      private String documentId = "";
      private boolean sawComplete;
      private boolean aborted;

      @Override
      public void onNext(ParseRequestChunk chunk) {
        if (aborted) return;
        if (documentId.isEmpty() && !chunk.getDocumentId().isEmpty()) {
          documentId = chunk.getDocumentId();
        }
        if (buffer.size() + (long) chunk.getData().size() > maxDocumentBytes) {
          aborted = true;
          counters.recordRejected();
          responses.onError(
              Status.RESOURCE_EXHAUSTED
                  .withDescription("document exceeds " + maxDocumentBytes + " bytes")
                  .asRuntimeException());
          return;
        }
        try {
          chunk.getData().writeTo(buffer);
        } catch (java.io.IOException impossible) {
          throw new IllegalStateException("in-memory buffer write failed", impossible);
        }
        if (chunk.getComplete()) sawComplete = true;
      }

      @Override
      public void onError(Throwable error) {
        aborted = true;
      }

      @Override
      public void onCompleted() {
        if (aborted) return;
        if (!sawComplete || buffer.size() == 0) {
          counters.recordRejected();
          responses.onError(
              Status.INVALID_ARGUMENT
                  .withDescription(buffer.size() == 0
                      ? "no document bytes received"
                      : "stream ended without a chunk marked complete")
                  .asRuntimeException());
          return;
        }
        byte[] bytes = buffer.toByteArray();
        String id = documentId;
        executor.execute(() -> parseNow(id, bytes, responses));
      }
    };
  }

  private void parseNow(String documentId, byte[] bytes, StreamObserver<ParseEvent> responses) {
    try {
      parseSlots.acquire();
    } catch (InterruptedException interrupt) {
      Thread.currentThread().interrupt();
      responses.onError(Status.UNAVAILABLE.withDescription("server shutting down")
          .asRuntimeException());
      return;
    }
    try {
      parser.parse(documentId, bytes, responses::onNext);
      responses.onCompleted();
      counters.recordParsed();
    } catch (UnsupportedFormatException unsupported) {
      counters.recordRejected();
      fail(responses, Status.UNIMPLEMENTED.withDescription(unsupported.getMessage()));
    } catch (InvalidDocumentException invalid) {
      counters.recordRejected();
      fail(responses, Status.INVALID_ARGUMENT.withDescription(invalid.getMessage()));
    } catch (Exception unexpected) {
      counters.recordFailed();
      fail(responses, Status.INTERNAL.withDescription("parser fault: " + unexpected.getMessage()));
    } catch (OutOfMemoryError exhausted) {
      // The document's allocations become unreachable as the stack unwinds,
      // so the call can still be failed; rethrown so the log shows it.
      counters.recordFailed();
      fail(responses, OUT_OF_MEMORY);
      throw exhausted;
    } catch (StackOverflowError tooDeep) {
      counters.recordFailed();
      fail(responses, TOO_DEEP);
    } catch (Throwable fault) {
      // Any other Error (a linkage failure, an assertion) must still close
      // the call: an open call leaves the client waiting for its deadline.
      counters.recordFailed();
      fail(responses, Status.INTERNAL
          .withDescription("parser fault: " + fault.getClass().getSimpleName()));
      if (fault instanceof VirtualMachineError machineError) throw machineError;
    } finally {
      parseSlots.release();
    }
  }

  /** Fails the call; a call the client already abandoned has nothing left to fail. */
  private static void fail(StreamObserver<ParseEvent> responses, Status status) {
    try {
      responses.onError(status.asRuntimeException());
    } catch (RuntimeException alreadyClosed) {
      // cancelled or closed underneath us
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
