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

  private final long maxDocumentBytes;
  private final int maxConcurrentParses;
  private final Semaphore parseSlots;
  private final ExecutorService executor;
  private final ParseCounters counters = new ParseCounters();

  public PoiParseServiceImpl(long maxDocumentBytes, int maxConcurrentParses,
                             ExecutorService executor) {
    this.maxDocumentBytes = maxDocumentBytes;
    this.maxConcurrentParses = maxConcurrentParses;
    this.parseSlots = new Semaphore(maxConcurrentParses);
    this.executor = executor;
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
      DocumentParser.parse(documentId, bytes, responses::onNext);
      responses.onCompleted();
      counters.recordParsed();
    } catch (UnsupportedFormatException unsupported) {
      counters.recordRejected();
      responses.onError(Status.UNIMPLEMENTED.withDescription(unsupported.getMessage())
          .asRuntimeException());
    } catch (InvalidDocumentException invalid) {
      counters.recordRejected();
      responses.onError(Status.INVALID_ARGUMENT.withDescription(invalid.getMessage())
          .asRuntimeException());
    } catch (Exception unexpected) {
      counters.recordFailed();
      responses.onError(Status.INTERNAL
          .withDescription("parser fault: " + unexpected.getMessage()).asRuntimeException());
    } finally {
      parseSlots.release();
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
