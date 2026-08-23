package ai.pipestream.grpoic.server;

import io.grpc.Grpc;
import io.grpc.InsecureServerCredentials;
import io.grpc.Server;
import io.grpc.protobuf.services.HealthStatusManager;
import io.grpc.protobuf.services.ProtoReflectionService;
import io.grpc.protobuf.services.ProtoReflectionServiceV1;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

/**
 * grPOIc: a diskless gRPC wrapper over Apache POI. Documents arrive as byte
 * streams, parse entirely in memory, and leave as typed structure events.
 * Nothing is written to disk and no external process is executed, ever.
 */
public final class GrPoicServer {

  // v1alpha reflection stays registered alongside v1: older grpcurl and
  // orchestrator builds still dial it, and dropping it would be wire-visible.
  @SuppressWarnings("deprecation")
  public static void main(String[] args) throws Exception {
    final int port = intFromEnv("GRPOIC_PORT", 50052, 1, 65535);
    final long maxDocumentBytes =
        intFromEnv("GRPOIC_MAX_DOCUMENT_MIB", 70, 1, 1024) * 1024L * 1024L;
    final int cores = Runtime.getRuntime().availableProcessors();
    final int maxConcurrent =
        intFromEnv("GRPOIC_MAX_CONCURRENT_PARSES", Math.max(2, cores), 1, 256);
    final int metricsInterval = intFromEnv("GRPOIC_METRICS_INTERVAL_SECONDS", 60, 0, 86400);

    var executor = Executors.newVirtualThreadPerTaskExecutor();
    PoiParseServiceImpl service =
        new PoiParseServiceImpl(maxDocumentBytes, maxConcurrent, executor);
    HealthStatusManager health = new HealthStatusManager();
    Server server =
        Grpc.newServerBuilderForPort(port, InsecureServerCredentials.create())
            // Allow single-chunk uploads of a full-size document plus framing.
            .maxInboundMessageSize((int) Math.min(Integer.MAX_VALUE, maxDocumentBytes + (1 << 20)))
            .addService(service)
            .addService(health.getHealthService())
            .addService(ProtoReflectionService.newInstance())
            .addService(ProtoReflectionServiceV1.newInstance())
            .build()
            .start();
    System.out.println("grPOIc " + PoiParseServiceImpl.SERVICE_VERSION + " listening on 0.0.0.0:"
        + port + " (POI " + org.apache.poi.Version.getVersion() + ", max "
        + (maxDocumentBytes >> 20) + " MiB, " + maxConcurrent + " concurrent parses)");

    ScheduledExecutorService metrics = startMetrics(service.counters(), metricsInterval);

    Runtime.getRuntime().addShutdownHook(new Thread(() -> {
      server.shutdown();
      try {
        if (!server.awaitTermination(30, TimeUnit.SECONDS)) server.shutdownNow();
      } catch (InterruptedException interrupt) {
        server.shutdownNow();
      }
      if (metrics != null) metrics.shutdownNow();
      executor.shutdown();
    }, "grpoic-shutdown"));
    server.awaitTermination();
  }

  private static ScheduledExecutorService startMetrics(ParseCounters counters, int intervalSeconds) {
    if (intervalSeconds <= 0) return null;
    ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor(task -> {
      Thread thread = new Thread(task, "grpoic-metrics");
      thread.setDaemon(true);
      return thread;
    });
    scheduler.scheduleAtFixedRate(
        () -> System.out.println("grPOIc metrics: " + counters.summary()),
        intervalSeconds, intervalSeconds, TimeUnit.SECONDS);
    return scheduler;
  }

  private static int intFromEnv(String name, int fallback, int min, int max) {
    String configured = System.getenv(name);
    if (configured == null || configured.isBlank()) return fallback;
    final int value;
    try {
      value = Integer.parseInt(configured.strip());
    } catch (NumberFormatException bad) {
      throw new IllegalArgumentException(name + " must be an integer, got: " + configured);
    }
    if (value < min || value > max) {
      throw new IllegalArgumentException(name + " must be in [" + min + ", " + max + "], got: "
          + value);
    }
    return value;
  }
}
