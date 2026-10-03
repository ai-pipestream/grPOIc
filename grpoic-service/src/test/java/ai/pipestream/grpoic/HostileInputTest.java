package ai.pipestream.grpoic;

import static ai.pipestream.grpoic.OoxmlFixtures.W;
import static ai.pipestream.grpoic.OoxmlFixtures.withPart;
import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Documents built to hurt the server: zip bombs, allocation claims, entity
 * tricks, and text amplification. Each is turned away with a status, and
 * none can make the server read a local file or allocate past its limits.
 */
class HostileInputTest {

  private static final long CAP = 4L * 1024 * 1024;
  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(CAP, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static byte[] docx() throws Exception {
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText("ordinary");
      document.write(out);
      return out.toByteArray();
    }
  }

  private static Status failure(ParseResult result) {
    assertThat(result.error()).isInstanceOf(StatusRuntimeException.class);
    return ((StatusRuntimeException) result.error()).getStatus();
  }

  @Test
  void zipBombEntryIsRejected() throws Exception {
    // 3 MB of blanks deflate to a few KB: far past POI's 1:100 ratio.
    byte[] bytes = withPart(docx(), "/word/document.xml",
        "<w:document xmlns:w=\"" + W + "\"><!--" + " ".repeat(3_000_000) + "--><w:body/>"
            + "</w:document>");
    assertThat(bytes.length).isLessThan(100_000);
    ParseResult result = harness.parse(bytes, "bomb", bytes.length);
    assertThat(failure(result).getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT);
  }

  @Test
  void allocationTableClaimCannotOutgrowTheCap() throws Exception {
    byte[] bytes;
    try (POIFSFileSystem container = new POIFSFileSystem();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      container.createDocument(new ByteArrayInputStream(new byte[64]), "WordDocument");
      container.writeFilesystem(out);
      bytes = out.toByteArray();
    }
    // Header field 0x2C: number of FAT sectors. 3000 sectors claim a 196 MB
    // file, which POI would allocate before noticing the file is 2 KB.
    ByteBuffer.wrap(bytes).order(ByteOrder.LITTLE_ENDIAN).putInt(0x2C, 3000);
    Status status = failure(harness.parse(bytes, "fat-claim", bytes.length));
    assertThat(status.getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT);
    assertThat(status.getDescription())
        .as("the allocation cap fires before the claimed table is allocated")
        .startsWith("malformed document: RecordFormatException");
  }

  @Test
  void externalEntitiesAreNeverResolved(@TempDir Path directory) throws Exception {
    Path secret = directory.resolve("secret.txt");
    Files.writeString(secret, "TOP-SECRET-MARKER");
    byte[] bytes = withPart(docx(), "/word/document.xml",
        "<?xml version=\"1.0\"?><!DOCTYPE w:document [<!ENTITY xxe SYSTEM \""
            + secret.toUri() + "\">]><w:document xmlns:w=\"" + W + "\"><w:body><w:p><w:r>"
            + "<w:t>&xxe;</w:t></w:r></w:p></w:body></w:document>");
    ParseResult result = harness.parse(bytes, "xxe", bytes.length);
    assertThat(result.events().toString()).doesNotContain("TOP-SECRET-MARKER");
    assertThat(String.valueOf(result.error())).doesNotContain("TOP-SECRET-MARKER");
    assertThat(failure(result).getCode())
        .as("a DOCTYPE in a part is refused outright")
        .isEqualTo(Status.Code.INVALID_ARGUMENT);
  }

  @Test
  void entityExpansionIsRefused() throws Exception {
    StringBuilder dtd = new StringBuilder("<!ENTITY a0 \"lol\">");
    for (int level = 1; level < 10; level++) {
      String ref = "&a" + (level - 1) + ";";
      dtd.append("<!ENTITY a").append(level).append(" \"").append(ref.repeat(10)).append("\">");
    }
    byte[] bytes = withPart(docx(), "/word/document.xml",
        "<?xml version=\"1.0\"?><!DOCTYPE w:document [" + dtd + "]><w:document xmlns:w=\""
            + W + "\"><w:body><w:p><w:r><w:t>&a9;</w:t></w:r></w:p></w:body></w:document>");
    assertThat(failure(harness.parse(bytes, "laughs", bytes.length)).getCode())
        .isEqualTo(Status.Code.INVALID_ARGUMENT);
  }

  @Test
  void sharedStringAmplificationHitsTheTextBudget() throws Exception {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      String repeated = "x".repeat(10_000);
      var sheet = workbook.createSheet("Amplified");
      for (int row = 0; row < 50; row++) {
        var cells = sheet.createRow(row);
        for (int column = 0; column < 20; column++) cells.createCell(column).setCellValue(repeated);
      }
      workbook.write(out);
      bytes = out.toByteArray();
    }
    assertThat(bytes.length).as("one shared string, a thousand references").isLessThan(50_000);
    long installed = ZipSecureFile.getMaxTextSize();
    ZipSecureFile.setMaxTextSize(1_000_000);
    try {
      Status status = failure(harness.parse(bytes, "amplified", bytes.length));
      assertThat(status.getCode()).isEqualTo(Status.Code.RESOURCE_EXHAUSTED);
      assertThat(status.getDescription()).contains("exceeds 1000000 characters");
    } finally {
      ZipSecureFile.setMaxTextSize(installed);
    }
  }
}
