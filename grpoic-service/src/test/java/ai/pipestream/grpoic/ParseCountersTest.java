package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;
import static org.awaitility.Awaitility.await;

import java.io.ByteArrayOutputStream;
import java.time.Duration;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/**
 * The lifetime counters behind the stdout metrics line. Parsing completes on
 * a virtual thread after the response stream closes, so counter observation
 * is genuinely asynchronous: Awaitility, never a fixed sleep.
 */
class ParseCountersTest {

  private static byte[] xlsx() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      workbook.createSheet("S").createRow(0).createCell(0).setCellValue(1.0);
      workbook.write(out);
      return out.toByteArray();
    }
  }

  @Test
  void countersTrackParsedAndRejectedOutcomes() throws Exception {
    try (ParseHarness harness = new ParseHarness(1024 * 1024, 2)) {
      byte[] good = xlsx();
      assertThat(harness.parseOk(good, "ok-1").status().getSheets()).isEqualTo(1);
      assertThat(harness.parse(new byte[64], "junk", 64).error())
          .as("garbage is rejected")
          .isNotNull();
      assertThat(harness.parse(new byte[2 * 1024 * 1024], "big", 2 * 1024 * 1024).error())
          .as("over-cap is rejected")
          .isNotNull();

      await().atMost(Duration.ofSeconds(5)).untilAsserted(() -> {
        assertThat(harness.service().counters().parsed()).isEqualTo(1);
        assertThat(harness.service().counters().rejected()).isEqualTo(2);
        assertThat(harness.service().counters().failed()).isZero();
      });
      assertThat(harness.service().counters().summary())
          .as("the stdout metrics line format is observable and pinned")
          .isEqualTo("docs{parsed=1,rejected=2,failed=0}");
    }
  }
}
