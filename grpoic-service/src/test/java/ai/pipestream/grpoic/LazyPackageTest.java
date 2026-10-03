package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.poi.v1.ParseEvent;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * OOXML packages are read in place from the upload buffer: parts inflate
 * only when a parser asks for them, a damaged central directory falls back
 * to reading local headers, and entry names get POI's own checks.
 */
class LazyPackageTest {

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static byte[] docx(String text) throws Exception {
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText(text);
      document.write(out);
      return out.toByteArray();
    }
  }

  private static String onlyParagraph(byte[] bytes, String id) throws Exception {
    var paragraphs = harness.parseOk(bytes, id).eventsOf(ParseEvent::hasParagraph);
    assertThat(paragraphs).hasSize(1);
    return paragraphs.get(0).getParagraph().getText();
  }

  @Test
  void partsNoParserReadsAreNeverInflated() throws Exception {
    // 20 MB of text from a small vocabulary: compresses about 12:1, well
    // inside the zip-bomb ratio, but inflates past the 12 MiB cap on an
    // in-memory entry at this server's 4 MiB document cap.
    String[] words = {"alpha", "beta", "gamma", "delta", "omega", "sigma", "kappa", "theta"};
    Random random = new Random(7);
    StringBuilder padding = new StringBuilder(20_000_000);
    while (padding.length() < 20_000_000) padding.append(words[random.nextInt(words.length)]);
    byte[] bytes;
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(docx("still here")));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      PackagePart part = container.createPart(
          PackagingURIHelper.createPartName("/word/media/padding.bin"), "application/octet-stream");
      try (OutputStream partOut = part.getOutputStream()) {
        partOut.write(padding.toString().getBytes(StandardCharsets.US_ASCII));
      }
      container.save(out);
      bytes = out.toByteArray();
    }
    assertThat(bytes.length).isLessThan(4 * 1024 * 1024);
    assertThat(onlyParagraph(bytes, "padded")).isEqualTo("still here");
  }

  @Test
  void damagedCentralDirectoryFallsBackToLocalHeaders() throws Exception {
    byte[] bytes = docx("found without a directory");
    // Point the end-of-central-directory record (PK 05 06) at offset 0, so a
    // random-access reader finds a local header where it expects the
    // directory; reading the local headers in order still works.
    for (int at = bytes.length - 22; at >= 0; at--) {
      if (bytes[at] == 'P' && bytes[at + 1] == 'K' && bytes[at + 2] == 5 && bytes[at + 3] == 6) {
        ByteBuffer.wrap(bytes).order(ByteOrder.LITTLE_ENDIAN).putInt(at + 16, 0);
        break;
      }
    }
    assertThat(onlyParagraph(bytes, "no-directory")).isEqualTo("found without a directory");
  }

  @Test
  void entryNamesDifferingOnlyByCaseAreRejected() throws Exception {
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    try (ZipInputStream in = new ZipInputStream(new ByteArrayInputStream(docx("original")));
         ZipOutputStream zip = new ZipOutputStream(out)) {
      for (ZipEntry entry = in.getNextEntry(); entry != null; entry = in.getNextEntry()) {
        byte[] content = in.readAllBytes();
        zip.putNextEntry(new ZipEntry(entry.getName()));
        zip.write(content);
        if (entry.getName().equals("word/document.xml")) {
          zip.putNextEntry(new ZipEntry("WORD/Document.xml"));
          zip.write(content);
        }
      }
    }
    byte[] bytes = out.toByteArray();
    var result = harness.parse(bytes, "shadowed", bytes.length);
    assertThat(result.error()).isInstanceOf(StatusRuntimeException.class);
    assertThat(((StatusRuntimeException) result.error()).getStatus().getCode())
        .isEqualTo(Status.Code.INVALID_ARGUMENT);
  }
}
