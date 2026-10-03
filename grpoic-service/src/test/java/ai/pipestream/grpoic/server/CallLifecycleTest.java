package ai.pipestream.grpoic.server;

import static org.assertj.core.api.Assertions.assertThat;
import static org.awaitility.Awaitility.await;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseRequestChunk;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.PoiParseServiceGrpc;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.ClientCallStreamObserver;
import io.grpc.stub.ClientResponseObserver;
import java.time.Duration;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

/**
 * A call's life as the transport sees it: admission before the upload is
 * read, backpressure on the way out, and cancellation, a stalled upload or
 * a client that stopped reading giving the parse slot back. The parser seam emits plain events so the
 * flow, not the document, is under test.
 */
class CallLifecycleTest {

  private ExecutorService executor;
  private Server server;
  private ManagedChannel channel;
  private PoiParseServiceImpl service;

  private void start(int slots, PoiParseServiceImpl.Parser parser, Duration idle)
      throws Exception {
    start(slots, parser, idle, PoiParseServiceImpl.UPLOAD_TIME_LIMIT);
  }

  private void start(int slots, PoiParseServiceImpl.Parser parser, Duration idle,
                     Duration limit) throws Exception {
    start(slots, parser, idle, limit, PoiParseServiceImpl.RESPONSE_STALL_LIMIT);
  }

  private void start(int slots, PoiParseServiceImpl.Parser parser, Duration idle,
                     Duration limit, Duration stall) throws Exception {
    executor = Executors.newVirtualThreadPerTaskExecutor();
    service = new PoiParseServiceImpl(1 << 20, slots, executor, parser, idle, limit, stall);
    String name = InProcessServerBuilder.generateName();
    server = InProcessServerBuilder.forName(name).directExecutor()
        .addService(service).build().start();
    channel = InProcessChannelBuilder.forName(name).directExecutor().build();
  }

  @AfterEach
  void stop() throws Exception {
    channel.shutdownNow();
    server.shutdownNow();
    executor.shutdown();
    channel.awaitTermination(5, TimeUnit.SECONDS);
  }

  private static ParseEvent event(int number) {
    return ParseEvent.newBuilder()
        .setStatus(ParseStatus.newBuilder().setParagraphs(number))
        .build();
  }

  /** A client that reads responses only when told to. */
  private final class Client implements ClientResponseObserver<ParseRequestChunk, ParseEvent> {
    final AtomicInteger received = new AtomicInteger();
    final AtomicReference<Throwable> failure = new AtomicReference<>();
    final CountDownLatch done = new CountDownLatch(1);
    final int initialRequests;
    ClientCallStreamObserver<ParseRequestChunk> call;

    Client(int initialRequests) {
      this.initialRequests = initialRequests;
    }

    Client upload(boolean complete) {
      PoiParseServiceGrpc.newStub(channel).parseDocument(this);
      call.onNext(ParseRequestChunk.newBuilder().setDocumentId("flow")
          .setData(ByteString.copyFromUtf8("bytes")).setComplete(complete).build());
      if (complete) call.onCompleted();
      return this;
    }

    /** Sends a chunk of {@code data} every 50 ms, never one marked complete. */
    Client trickle(ByteString data) {
      PoiParseServiceGrpc.newStub(channel).parseDocument(this);
      Thread.ofVirtual().start(() -> {
        try {
          while (done.getCount() > 0) {
            call.onNext(ParseRequestChunk.newBuilder().setDocumentId("drip").setData(data).build());
            Thread.sleep(50);
          }
        } catch (InterruptedException | RuntimeException stopped) {
          // the call ended underneath the sender
        }
      });
      return this;
    }

    @Override
    public void beforeStart(ClientCallStreamObserver<ParseRequestChunk> requestStream) {
      call = requestStream;
      if (initialRequests >= 0) requestStream.disableAutoRequestWithInitial(initialRequests);
    }

    @Override
    public void onNext(ParseEvent event) {
      received.incrementAndGet();
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
  }

  @Test
  void slowReaderStallsTheParseInsteadOfQueueingEvents() throws Exception {
    AtomicInteger emitted = new AtomicInteger();
    start(2, (id, data, options, emit) -> {
      for (int number = 0; number < 100; number++) {
        emit.accept(event(number));
        emitted.incrementAndGet();
      }
    }, PoiParseServiceImpl.UPLOAD_IDLE_TIMEOUT);

    Client client = new Client(2).upload(true);
    await().atMost(Duration.ofSeconds(5)).until(() -> client.received.get() == 2);
    Thread.sleep(300);
    assertThat(emitted.get())
        .as("the parse waits for the reader instead of running ahead of it")
        .isLessThanOrEqualTo(3);

    client.call.request(1000);
    assertThat(client.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(client.failure.get()).isNull();
    assertThat(client.received.get()).isEqualTo(100);
  }

  @Test
  void cancellingAStalledCallFreesItsSlot() throws Exception {
    AtomicInteger calls = new AtomicInteger();
    start(1, (id, data, options, emit) -> {
      if (calls.getAndIncrement() == 0) {
        for (int number = 0; ; number++) emit.accept(event(number));
      }
      emit.accept(event(1));
    }, PoiParseServiceImpl.UPLOAD_IDLE_TIMEOUT);

    Client first = new Client(1).upload(true);
    await().atMost(Duration.ofSeconds(5)).until(() -> first.received.get() == 1);
    first.call.cancel("client gave up", null);

    // One slot only: the second call can run only if the first gave it back.
    Client second = new Client(-1).upload(true);
    assertThat(second.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(second.failure.get()).isNull();
    await().atMost(Duration.ofSeconds(5)).untilAsserted(() ->
        assertThat(service.counters().summary())
            .as("a cancelled call is neither parsed nor failed")
            .isEqualTo("docs{parsed=1,rejected=0,failed=0}"));
  }

  @Test
  void queuedCallDoesNotReadItsUploadUntilAdmitted() throws Exception {
    CountDownLatch release = new CountDownLatch(1);
    AtomicInteger calls = new AtomicInteger();
    start(1, (id, data, options, emit) -> {
      if (calls.getAndIncrement() == 0) {
        try {
          release.await();
        } catch (InterruptedException interrupt) {
          Thread.currentThread().interrupt();
        }
      }
      emit.accept(event(1));
    }, PoiParseServiceImpl.UPLOAD_IDLE_TIMEOUT);

    Client holder = new Client(-1).upload(true);
    await().atMost(Duration.ofSeconds(5)).until(() -> calls.get() == 1);
    Client queued = new Client(-1).upload(true);
    Thread.sleep(300);
    assertThat(queued.call.isReady())
        .as("no slot, no request: the queued upload stays with the client")
        .isFalse();

    release.countDown();
    assertThat(holder.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(queued.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(queued.failure.get()).isNull();
    assertThat(queued.received.get()).isEqualTo(1);
  }

  @Test
  void stalledUploadGivesItsSlotBack() throws Exception {
    start(1, (id, data, options, emit) -> emit.accept(event(1)), Duration.ofMillis(200));

    Client stalled = new Client(-1).upload(false);
    assertThat(stalled.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(((StatusRuntimeException) stalled.failure.get()).getStatus().getCode())
        .isEqualTo(Status.Code.DEADLINE_EXCEEDED);

    Client next = new Client(-1).upload(true);
    assertThat(next.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(next.failure.get()).isNull();
    // The parsed count lands just after the call completes.
    await().atMost(Duration.ofSeconds(5)).untilAsserted(() ->
        assertThat(service.counters().summary()).isEqualTo("docs{parsed=1,rejected=1,failed=0}"));
  }

  @Test
  void tricklingUploadIsCutOffAtTheTimeLimit() throws Exception {
    start(1, (id, data, options, emit) -> emit.accept(event(1)), Duration.ofMillis(300),
        Duration.ofSeconds(1));

    // A byte every 50 ms never trips the idle timeout; only the total limit ends it.
    Client drip = new Client(-1).trickle(ByteString.copyFromUtf8("x"));
    assertThat(drip.done.await(10, TimeUnit.SECONDS)).isTrue();
    Status status = ((StatusRuntimeException) drip.failure.get()).getStatus();
    assertThat(status.getCode()).isEqualTo(Status.Code.DEADLINE_EXCEEDED);
    assertThat(status.getDescription()).startsWith("upload took longer than");

    Client next = new Client(-1).upload(true);
    assertThat(next.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(next.failure.get()).isNull();
  }

  @Test
  void clientThatStopsReadingGivesItsSlotBack() throws Exception {
    AtomicInteger calls = new AtomicInteger();
    start(1, (id, data, options, emit) -> {
      if (calls.getAndIncrement() == 0) {
        for (int number = 0; ; number++) emit.accept(event(number));
      }
      emit.accept(event(1));
    }, PoiParseServiceImpl.UPLOAD_IDLE_TIMEOUT, PoiParseServiceImpl.UPLOAD_TIME_LIMIT,
        Duration.ofMillis(300));

    // Uploads, reads one event, then never reads again and never cancels.
    Client silent = new Client(1).upload(true);
    await().atMost(Duration.ofSeconds(5)).until(() -> silent.received.get() == 1);

    // One slot only: the second call can run only if the first gave it back.
    Client next = new Client(-1).upload(true);
    assertThat(next.done.await(10, TimeUnit.SECONDS)).isTrue();
    assertThat(next.failure.get()).isNull();

    // Once it reads again, the first client finds its call failed.
    silent.call.request(1000);
    assertThat(silent.done.await(10, TimeUnit.SECONDS)).isTrue();
    Status status = ((StatusRuntimeException) silent.failure.get()).getStatus();
    assertThat(status.getCode()).isEqualTo(Status.Code.DEADLINE_EXCEEDED);
    assertThat(status.getDescription()).startsWith("client read no response for");
    await().atMost(Duration.ofSeconds(5)).untilAsserted(() ->
        assertThat(service.counters().summary()).isEqualTo("docs{parsed=1,rejected=1,failed=0}"));
  }

  @Test
  void emptyChunksDoNotCountAsProgress() throws Exception {
    start(1, (id, data, options, emit) -> emit.accept(event(1)), Duration.ofMillis(300),
        Duration.ofMinutes(5));

    Client empty = new Client(-1).trickle(ByteString.EMPTY);
    assertThat(empty.done.await(10, TimeUnit.SECONDS)).isTrue();
    Status status = ((StatusRuntimeException) empty.failure.get()).getStatus();
    assertThat(status.getCode()).isEqualTo(Status.Code.DEADLINE_EXCEEDED);
    assertThat(status.getDescription()).startsWith("no upload data for");
  }
}
