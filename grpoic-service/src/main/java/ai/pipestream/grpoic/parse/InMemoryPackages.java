package ai.pipestream.grpoic.parse;

import com.google.protobuf.ByteString;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipArchiveThresholdInputStream;
import org.apache.poi.openxml4j.util.ZipFileZipEntrySource;

/**
 * Opens an OOXML package from the upload buffer without unpacking it.
 *
 * <p><b>Why not {@code OPCPackage.open(InputStream)}.</b> The stream form
 * inflates every entry into its own byte array before the first part is
 * parsed, so a package costs its full decompressed size in heap on top of
 * the upload, and one entry past POI's in-memory entry cap fails the whole
 * document. Reading the central directory instead gives random access:
 * each part is inflated as a stream, only when a parser asks for it, and
 * parts nothing reads are never inflated at all.
 *
 * <p><b>Same guards.</b> Every entry still inflates through POI's threshold
 * stream (the ratio and size limits {@link PoiLimits} sets), entry names
 * get the checks POI's own zip file applies (no empty names, no names that
 * differ only by case), and the package's entry count is checked by POI.
 *
 * <p><b>Damaged directories.</b> When the central directory is unreadable,
 * this falls back to POI's stream reading of the local headers, as POI
 * itself does for files, with each entry held in memory under the
 * allocation cap.
 */
final class InMemoryPackages {

  private InMemoryPackages() {}

  static OPCPackage open(ByteString data) throws IOException, InvalidFormatException {
    ZipFile zip;
    try {
      zip = ZipFile.builder().setSeekableByteChannel(new ByteStringChannel(data)).get();
    } catch (IOException unreadableDirectory) {
      return OPCPackage.open(data.newInput());
    }
    try {
      validateNames(zip);
      return OPCPackage.open(new ThresholdEntrySource(zip));
    } catch (IOException | InvalidFormatException | RuntimeException error) {
      zip.close();
      throw error;
    }
  }

  private static void validateNames(ZipFile zip) throws IOException {
    Set<String> seen = new HashSet<>();
    Enumeration<ZipArchiveEntry> entries = zip.getEntries();
    while (entries.hasMoreElements()) {
      String name = entries.nextElement().getName();
      if (name.isEmpty()) throw new IOException("zip entry with an empty name");
      if (!seen.add(name.toLowerCase(Locale.ROOT))) {
        throw new IOException("more than one zip entry named " + name);
      }
    }
  }

  /** Inflates every entry through POI's zip-bomb guard. */
  private static final class ThresholdEntrySource extends ZipFileZipEntrySource {
    ThresholdEntrySource(ZipFile zip) {
      super(zip);
    }

    @Override
    public InputStream getInputStream(ZipArchiveEntry entry) throws IOException {
      return new ZipArchiveThresholdInputStream(super.getInputStream(entry));
    }
  }
}
