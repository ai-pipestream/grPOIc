package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.IOException;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.util.ZipArchiveFakeEntry;
import org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.RecordFormatException;
import org.apache.poi.util.TempFile;
import org.junit.jupiter.api.Test;

/**
 * The JVM-wide POI limits are explicit, derived from the document cap, and
 * POI is refused every way of writing to disk.
 */
class PoiLimitsTest {

  private static final long MIB = 1L << 20;

  @Test
  void limitsFollowTheDocumentCap() {
    PoiLimits.install(70 * MIB);
    assertThat(ZipSecureFile.getMinInflateRatio()).isEqualTo(0.01);
    assertThat(ZipSecureFile.getMaxEntrySize()).isEqualTo(16 * 70 * MIB);
    assertThat(ZipSecureFile.getMaxFileCount()).isEqualTo(10_000);
    assertThat(ZipSecureFile.getMaxTextSize()).isEqualTo(4 * 70 * MIB);
    assertThat(IOUtils.getByteArrayMaxOverride()).isEqualTo((int) (78 * MIB));
    assertThat(ZipArchiveFakeEntry.getMaxEntrySize()).isEqualTo((int) (78 * MIB));

    PoiLimits.install(1 * MIB);
    assertThat(ZipSecureFile.getMaxEntrySize()).as("floor").isEqualTo(256 * MIB);
    assertThat(ZipSecureFile.getMaxTextSize()).as("floor").isEqualTo(10 * MIB);
    assertThat(IOUtils.getByteArrayMaxOverride()).isEqualTo((int) (9 * MIB));

    PoiLimits.install(1024 * MIB);
    assertThat(ZipSecureFile.getMaxEntrySize()).as("POI's 4 GiB ceiling").isEqualTo(0xFFFFFFFFL);
  }

  @Test
  void noAllocationMayExceedTheCap() {
    PoiLimits.install(4 * MIB);
    assertThatThrownBy(() -> IOUtils.safelyAllocateCheck(12 * MIB + 1, Integer.MAX_VALUE))
        .as("the override replaces POI's per-record ceilings, including its 250 MB ones")
        .isInstanceOf(RecordFormatException.class);
  }

  @Test
  void poiCannotWriteTemporaryFiles() {
    PoiLimits.install(4 * MIB);
    assertThat(ZipInputStreamZipEntrySource.getThresholdBytesForTempFiles()).isEqualTo(-1);
    assertThat(ZipPackage.useTempFilePackageParts()).isFalse();
    assertThatThrownBy(() -> TempFile.createTempFile("poi", ".tmp"))
        .isInstanceOf(IOException.class)
        .hasMessageContaining("diskless");
    assertThatThrownBy(() -> TempFile.createTempDirectory("poi"))
        .isInstanceOf(IOException.class)
        .hasMessageContaining("diskless");
  }
}
