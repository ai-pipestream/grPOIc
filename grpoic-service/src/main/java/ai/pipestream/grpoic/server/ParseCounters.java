package ai.pipestream.grpoic.server;

import java.util.concurrent.atomic.AtomicLong;

/**
 * Lifetime document counters: parsed to completion, rejected before or during
 * validation (bad bytes, over cap, unsupported format), and failed inside the
 * parser. The summary line is the wire-stable stdout metrics format.
 */
public final class ParseCounters {

  private final AtomicLong parsed = new AtomicLong();
  private final AtomicLong rejected = new AtomicLong();
  private final AtomicLong failed = new AtomicLong();

  void recordParsed() {
    parsed.incrementAndGet();
  }

  void recordRejected() {
    rejected.incrementAndGet();
  }

  void recordFailed() {
    failed.incrementAndGet();
  }

  public long parsed() {
    return parsed.get();
  }

  public long rejected() {
    return rejected.get();
  }

  public long failed() {
    return failed.get();
  }

  /** The stdout metrics line body; the format is observable and must not drift. */
  public String summary() {
    return "docs{parsed=" + parsed.get() + ",rejected=" + rejected.get()
        + ",failed=" + failed.get() + "}";
  }
}
