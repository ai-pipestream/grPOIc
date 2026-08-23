package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.server.PoiParseServiceImpl;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

/**
 * One in-process grPOIc rig per instance: real service, real transport, and a
 * synchronous round-trip helper. Kept out of the test classes so every suite
 * exercises the identical wire path.
 */
final class ParseHarness implements AutoCloseable {

  private final Server server;
  private final ManagedChannel channel;
  private final ExecutorService executor;
  private final PoiParseServiceImpl service;

  ParseHarness(long maxDocumentBytes, int maxConcurrentParses) throws IOException {
    executor = Executors.newVirtualThreadPerTaskExecutor();
    service = new PoiParseServiceImpl(maxDocumentBytes, maxConcurrentParses, executor);
    String name = InProcessServerBuilder.generateName();
    server = InProcessServerBuilder.forName(name).directExecutor()
        .addService(service).build().start();
    channel = InProcessChannelBuilder.forName(name).directExecutor().build();
  }

  PoiParseServiceImpl service() {
    return service;
  }

  PoiParseServiceGrpc.PoiParseServiceBlockingStub blockingStub() {
    return PoiParseServiceGrpc.newBlockingStub(channel);
  }

  record ParseResult(List<ParseEvent> events, Throwable error) {
    ParseStatus status() {
      ParseEvent last = events.get(events.size() - 1);
      assertThat(last.hasStatus()).as("last event must be the status").isTrue();
      return last.getStatus();
    }

    List<ParseEvent> eventsOf(java.util.function.Predicate<ParseEvent> kind) {
      return events.stream().filter(kind).toList();
    }
  }

  ParseResult parse(byte[] bytes, String documentId, int chunkSize) throws InterruptedException {
    List<ParseEvent> events = new ArrayList<>();
    AtomicReference<Throwable> failure = new AtomicReference<>();
    CountDownLatch done = new CountDownLatch(1);
    StreamObserver<ParseRequestChunk> requests =
        PoiParseServiceGrpc.newStub(channel).parseDocument(new StreamObserver<>() {
          @Override
          public void onNext(ParseEvent event) {
            synchronized (events) {
              events.add(event);
            }
          }

          @Override
          public void onError(Throwable error) {
            failure.set(error);
            done.countDown();
          }

          @Override
          public void onCompleted() {
            done.countDown();
          }
        });
    for (int offset = 0; offset < bytes.length; offset += chunkSize) {
      int end = Math.min(bytes.length, offset + chunkSize);
      ParseRequestChunk.Builder chunk = ParseRequestChunk.newBuilder()
          .setData(ByteString.copyFrom(bytes, offset, end - offset))
          .setComplete(end == bytes.length);
      if (offset == 0) chunk.setDocumentId(documentId);
      requests.onNext(chunk.build());
    }
    requests.onCompleted();
    assertThat(done.await(30, TimeUnit.SECONDS)).as("parse timed out").isTrue();
    return new ParseResult(events, failure.get());
  }

  ParseResult parseOk(byte[] bytes, String documentId) throws InterruptedException {
    ParseResult result = parse(bytes, documentId, bytes.length);
    assertThat(result.error()).as("parse must succeed").isNull();
    assertThat(result.events().get(0).hasDocumentInfo()).as("first event is document info")
        .isTrue();
    return result;
  }

  @Override
  public void close() throws InterruptedException {
    channel.shutdownNow();
    server.shutdownNow();
    executor.shutdown();
    assertThat(channel.awaitTermination(5, TimeUnit.SECONDS)).as("channel drain").isTrue();
  }
}
