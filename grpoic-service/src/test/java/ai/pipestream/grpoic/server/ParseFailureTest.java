package ai.pipestream.grpoic.server;

import static org.assertj.core.api.Assertions.assertThat;
import static org.awaitility.Awaitility.await;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import java.time.Duration;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/**
 * A parse that dies with an Error (out of memory, stack overflow, a linkage
 * failure) must still close the call with a status. Before, the Error
 * escaped to the virtual thread's uncaught handler and the client waited
 * for its deadline. The parser seam throws what no document reliably can.
 */
class ParseFailureTest {

  private record Outcome(Status status, ParseCounters counters) {}

  private static Outcome parseWith(PoiParseServiceImpl.Parser parser) throws Exception {
    ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
    PoiParseServiceImpl service = new PoiParseServiceImpl(1 << 20, 2, executor, parser);
    String name = InProcessServerBuilder.generateName();
    Server server = InProcessServerBuilder.forName(name).directExecutor()
        .addService(service).build().start();
    ManagedChannel channel = InProcessChannelBuilder.forName(name).directExecutor().build();
    try {
      AtomicReference<Throwable> failure = new AtomicReference<>();
      CountDownLatch done = new CountDownLatch(1);
      StreamObserver<ParseRequestChunk> requests =
          PoiParseServiceGrpc.newStub(channel).parseDocument(new StreamObserver<>() {
            @Override
            public void onNext(ParseEvent event) {}

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
      requests.onNext(ParseRequestChunk.newBuilder().setDocumentId("doomed")
          .setData(ByteString.copyFromUtf8("any bytes")).setComplete(true).build());
      requests.onCompleted();
      assertThat(done.await(10, TimeUnit.SECONDS))
          .as("the call must close instead of hanging until a deadline")
          .isTrue();
      assertThat(failure.get()).isInstanceOf(StatusRuntimeException.class);
      Status status = ((StatusRuntimeException) failure.get()).getStatus();
      // The service rethrows an OutOfMemoryError after failing the call, on
      // purpose; Awaitility would otherwise catch that uncaught throw and
      // fail this wait with it, aborting the test run, whenever the rethrow
      // lands after the wait has started.
      await().atMost(Duration.ofSeconds(5)).dontCatchUncaughtExceptions()
          .untilAsserted(() -> assertThat(service.counters().failed()).isEqualTo(1));
      return new Outcome(status, service.counters());
    } finally {
      channel.shutdownNow();
      server.shutdownNow();
      executor.shutdown();
    }
  }

  @Test
  void outOfMemoryFailsTheCallAsResourceExhausted() throws Exception {
    Outcome outcome = parseWith((id, bytes, options, emit) -> {
      throw new OutOfMemoryError("Java heap space");
    });
    assertThat(outcome.status().getCode()).isEqualTo(Status.Code.RESOURCE_EXHAUSTED);
    assertThat(outcome.counters().summary()).isEqualTo("docs{parsed=0,rejected=0,failed=1}");
  }

  @Test
  void realStackOverflowOnTheParseThreadFailsTheCall() throws Exception {
    Outcome outcome = parseWith((id, bytes, options, emit) -> recurse(0));
    assertThat(outcome.status().getCode()).isEqualTo(Status.Code.RESOURCE_EXHAUSTED);
    assertThat(outcome.status().getDescription()).contains("nests too deeply");
  }

  @Test
  void otherErrorsFailTheCallAsInternal() throws Exception {
    Outcome outcome = parseWith((id, bytes, options, emit) -> {
      throw new NoClassDefFoundError("org/example/Missing");
    });
    assertThat(outcome.status().getCode()).isEqualTo(Status.Code.INTERNAL);
    assertThat(outcome.status().getDescription()).isEqualTo("parser fault: NoClassDefFoundError");
  }

  @Test
  void serverFaultDescriptionsNameTheTypeWithoutItsMessage() throws Exception {
    Outcome outcome = parseWith((id, bytes, options, emit) -> {
      throw new IllegalStateException("cell text: Quarterly salaries for J. Doe");
    });
    assertThat(outcome.status().getCode()).isEqualTo(Status.Code.INTERNAL);
    assertThat(outcome.status().getDescription())
        .as("a message may quote the document and must not leave in the status")
        .isEqualTo("parser fault: IllegalStateException");
  }

  private static int recurse(int depth) {
    return recurse(depth + 1) + 1;
  }
}
