package ai.pipestream.grpoic;

import static ai.pipestream.grpoic.OoxmlFixtures.W;
import static ai.pipestream.grpoic.OoxmlFixtures.withPart;
import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetRow;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Broken input is the client's problem and says so: a corrupt document maps
 * to INVALID_ARGUMENT (counted as rejected), a pre-97 format to
 * UNIMPLEMENTED, an encrypted one to FAILED_PRECONDITION, and a single
 * broken cell or property part costs only itself, with a warning, instead
 * of failing the whole document as INTERNAL.
 */
class MalformedInputTest {

  static final String SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static Status failure(byte[] bytes, String id) throws Exception {
    ParseResult result = harness.parse(bytes, id, bytes.length);
    assertThat(result.error()).isInstanceOf(StatusRuntimeException.class);
    return ((StatusRuntimeException) result.error()).getStatus();
  }

  private static byte[] ole2(String... streams) throws Exception {
    try (POIFSFileSystem container = new POIFSFileSystem();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      for (int index = 0; index < streams.length; index += 2) {
        container.createDocument(
            new ByteArrayInputStream(streams[index + 1].getBytes()), streams[index]);
      }
      container.writeFilesystem(out);
      return out.toByteArray();
    }
  }

  private static byte[] ole2WordStream(byte[] wordDocument) throws Exception {
    try (POIFSFileSystem container = new POIFSFileSystem();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      container.createDocument(new ByteArrayInputStream(wordDocument), "WordDocument");
      container.writeFilesystem(out);
      return out.toByteArray();
    }
  }

  private static byte[] blankDocx() throws Exception {
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText("body survives");
      document.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] xlsxWithSheet(String sheetData) throws Exception {
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      workbook.createSheet("Data");
      workbook.write(out);
      base = out.toByteArray();
    }
    return withPart(base, "/xl/worksheets/sheet1.xml",
        "<worksheet xmlns=\"" + SHEET_NS + "\"><sheetData>" + sheetData
            + "</sheetData></worksheet>");
  }

  @Test
  void corruptBodyXmlIsInvalidArgument() throws Exception {
    byte[] bytes = withPart(blankDocx(), "/word/document.xml",
        "<w:document xmlns:w=\"" + W + "\"><w:body><w:p><w:r><w:t>never closed");
    Status status = failure(bytes, "corrupt-xml");
    assertThat(status.getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT);
    assertThat(status.getDescription()).startsWith("malformed document");
  }

  @Test
  void garbledLegacyWordStreamIsInvalidArgument() throws Exception {
    byte[] stream = new byte[4096];
    for (int index = 0; index < stream.length; index++) stream[index] = (byte) (index * 131 + 7);
    stream[0] = (byte) 0xEC; // FIB magic 0xA5EC, Word 97 nFib
    stream[1] = (byte) 0xA5;
    stream[2] = (byte) 0xC1;
    stream[3] = 0;
    assertThat(failure(ole2WordStream(stream), "garbled-doc").getCode())
        .as("an out-of-bounds read inside POI is the document's fault, not a parser fault")
        .isEqualTo(Status.Code.INVALID_ARGUMENT);
  }

  @Test
  void preWord97DocumentIsUnimplemented() throws Exception {
    byte[] stream = new byte[4096];
    stream[0] = (byte) 0xEC;
    stream[1] = (byte) 0xA5;
    stream[2] = 101; // Word 6.0
    Status status = failure(ole2WordStream(stream), "word6");
    assertThat(status.getCode()).isEqualTo(Status.Code.UNIMPLEMENTED);
    assertThat(status.getDescription()).startsWith("pre-97 office format");
  }

  @Test
  void encryptedOoxmlIsFailedPrecondition() throws Exception {
    byte[] bytes = ole2("EncryptionInfo", "agile descriptor", "EncryptedPackage", "ciphertext");
    Status status = failure(bytes, "encrypted");
    assertThat(status.getCode()).isEqualTo(Status.Code.FAILED_PRECONDITION);
    assertThat(status.getDescription()).contains("encrypted");
  }

  @Test
  void unreadableCellsCostOnlyThemselves() throws Exception {
    ParseResult result = harness.parseOk(xlsxWithSheet(
        "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#BOGUS!</v></c><c r=\"B1\"><v>7</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>not a number</v></c><c r=\"B2\" t=\"b\"><v>1</v></c></row>"),
        "bad-cells");
    var rows = result.eventsOf(ParseEvent::hasSheet).get(0).getSheet().getRowsList();
    assertThat(rows).extracting(SheetRow::getCellsCount).containsExactly(1, 1);
    assertThat(rows.get(0).getCells(0).getNumber()).isEqualTo(7.0);
    assertThat(rows.get(1).getCells(0).getBoolean()).isTrue();
    ParseStatus status = result.status();
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(status.getWarningsList())
        .anyMatch(warning -> warning.startsWith("sheet 'Data' cell A1 skipped"))
        .anyMatch(warning -> warning.startsWith("sheet 'Data' cell A2 skipped"));
  }

  @Test
  void brokenPropertyPartCostsOnlyTheMetadata() throws Exception {
    byte[] bytes = withPart(blankDocx(), "/docProps/app.xml", "<Properties><never-closed>");
    ParseResult result = harness.parseOk(bytes, "bad-props");
    assertThat(result.events().get(0).getDocumentInfo().getMetadata().getTailCount()).isZero();
    assertThat(result.eventsOf(ParseEvent::hasParagraph))
        .extracting(event -> event.getParagraph().getText())
        .containsExactly("body survives");
    assertThat(result.status().getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(result.status().getWarningsList())
        .anyMatch(warning -> warning.startsWith("document metadata skipped"));
  }
}
