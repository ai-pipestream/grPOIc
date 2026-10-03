package ai.pipestream.grpoic.parse;

import java.io.File;
import java.io.IOException;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.util.ZipArchiveFakeEntry;
import org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.TempFile;
import org.apache.poi.util.TempFileCreationStrategy;

/**
 * POI's process-wide safety limits, set explicitly from the document cap
 * instead of left at library defaults that assume a trusted desktop file.
 *
 * <p><b>Zip bombs.</b> Every OOXML entry is read through POI's threshold
 * stream: an entry may not inflate past {@link #maxEntryBytes} or below
 * POI's 1:100 compression ratio, and a package may not hold more than
 * {@link #MAX_FILE_COUNT} entries.
 *
 * <p><b>Allocations.</b> POI sizes many arrays from length fields inside
 * the file (an OLE2 header can claim a 250 MB allocation table in a few
 * hundred bytes). The byte-array override caps every such allocation at the
 * document cap plus slack, so no length field can ask for more memory than
 * the upload itself could justify.
 *
 * <p><b>Text.</b> {@link ZipSecureFile#setMaxTextSize} bounds the characters
 * of text a single document may yield. Spreadsheets enforce it per cell:
 * one shared string referenced by a million cells is the one way output can
 * outgrow its input many times over.
 *
 * <p><b>No temporary files.</b> The service is diskless by doctrine. POI's
 * spill-to-disk switches are pinned off and its temporary-file strategy
 * refuses, so any code path that would touch disk fails loudly instead of
 * writing document bytes to it.
 */
public final class PoiLimits {

  /** POI's own default, set explicitly: an entry may compress at most 100:1. */
  static final double MIN_INFLATE_RATIO = 0.01;
  /** Entries per package; POI's default of 1000 rejects large legitimate decks. */
  static final long MAX_FILE_COUNT = 10_000;

  private static final long MIB = 1L << 20;

  private PoiLimits() {}

  /**
   * Installs the limits for a server whose documents are at most
   * {@code maxDocumentBytes}. The settings are JVM-wide; installing again
   * replaces them.
   */
  public static synchronized void install(long maxDocumentBytes) {
    ZipSecureFile.setMinInflateRatio(MIN_INFLATE_RATIO);
    ZipSecureFile.setMaxEntrySize(maxEntryBytes(maxDocumentBytes));
    ZipSecureFile.setMaxFileCount(MAX_FILE_COUNT);
    ZipSecureFile.setMaxTextSize(maxTextChars(maxDocumentBytes));
    int arrayCap = maxAllocationBytes(maxDocumentBytes);
    IOUtils.setByteArrayMaxOverride(arrayCap);
    ZipArchiveFakeEntry.setMaxEntrySize(arrayCap);
    ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(-1);
    ZipPackage.setUseTempFilePackageParts(false);
    TempFile.setTempFileCreationStrategy(REFUSE_TEMP_FILES);
  }

  /**
   * The largest inflated size of one package entry: 16 times the document
   * cap (1.1 GiB at the default 70 MiB), at least 256 MiB, below POI's 4 GiB
   * ceiling.
   */
  static long maxEntryBytes(long maxDocumentBytes) {
    return Math.min(Math.max(16 * maxDocumentBytes, 256 * MIB), 0xFFFFFFFFL);
  }

  /** Characters of text one document may yield: four per byte of the cap. */
  static long maxTextChars(long maxDocumentBytes) {
    return Math.max(4 * maxDocumentBytes, 10 * MIB);
  }

  /**
   * The largest array POI may allocate: the document cap plus 8 MiB, since
   * no legitimate structure exceeds the file holding it (an OLE2 allocation
   * table rounds the file size up by at most one 4 MiB block).
   */
  static int maxAllocationBytes(long maxDocumentBytes) {
    return (int) Math.min(maxDocumentBytes + 8 * MIB, Integer.MAX_VALUE - 8);
  }

  private static final TempFileCreationStrategy REFUSE_TEMP_FILES =
      new TempFileCreationStrategy() {
        @Override
        public File createTempFile(String prefix, String suffix) throws IOException {
          throw new IOException("grPOIc is diskless: refused a temporary file " + prefix + suffix);
        }

        @Override
        public File createTempDirectory(String prefix) throws IOException {
          throw new IOException("grPOIc is diskless: refused a temporary directory " + prefix);
        }
      };
}
