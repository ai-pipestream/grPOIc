package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.GetServiceInfoRequest;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetCell;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.time.Instant;
import java.util.Arrays;
import java.util.Date;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Pins the extraction seams Apache POI has historically moved between
 * releases: whitespace in paragraph text, merged-region spans, cell
 * formatting, cached formula errors, date storage, placeholder routing and
 * metadata property mapping. Every expected value is exact; a POI upgrade
 * that reshapes any of these outputs is wire-visible and must fail here
 * before it reaches a client.
 */
class PoiSeamTest {

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  // --- DOCX text shape ----------------------------------------------------

  @Test
  void paragraphTextPreservesRunBoundariesAndTabs() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      XWPFParagraph paragraph = document.createParagraph();
      var first = paragraph.createRun();
      first.setText("Before");
      first.addTab();
      paragraph.createRun().setText("after  two spaces");
      document.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "runs");
    var paragraphs = result.eventsOf(ParseEvent::hasParagraph);
    assertThat(paragraphs).hasSize(1);
    assertThat(paragraphs.get(0).getParagraph().getText())
        .as("run concatenation, the literal tab and interior spaces are wire-visible; "
            + "POI must not reshape them")
        .isEqualTo("Before\tafter  two spaces");
  }

  @Test
  void blankParagraphsAreSkippedNotEmittedEmpty() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText("first");
      document.createParagraph(); // empty separator paragraph
      document.createParagraph().createRun().setText("second");
      document.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "blanks");
    assertThat(result.eventsOf(ParseEvent::hasParagraph))
        .extracting(event -> event.getParagraph().getText())
        .containsExactly("first", "second");
    assertThat(result.status().getParagraphs()).isEqualTo(2);
  }

  @Test
  void mergedTableCellsCarryColSpanOnTheAnchor() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      XWPFTable table = document.createTable(2, 3);
      var anchor = table.getRow(0).getCell(0);
      anchor.setText("spans two");
      anchor.getCTTc().addNewTcPr().addNewGridSpan().setVal(BigInteger.TWO);
      table.getRow(0).getCell(1).setText("third");
      table.getRow(1).getCell(0).setText("a");
      table.getRow(1).getCell(1).setText("b");
      table.getRow(1).getCell(2).setText("c");
      document.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "merge");
    var table = result.eventsOf(ParseEvent::hasTable).get(0).getTable();
    var anchor = table.getRows(0).getCells(0);
    assertThat(anchor.getText()).isEqualTo("spans two");
    assertThat(anchor.getColSpan()).as("gridSpan lands on the anchor cell").isEqualTo(2);
    assertThat(table.getRows(0).getCells(1).getColSpan()).isEqualTo(1);
    assertThat(table.getRows(1).getCellsCount()).isEqualTo(3);
  }

  @Test
  void emptyDocumentYieldsOnlyInfoAndOkStatus() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "empty");
    assertThat(result.events()).hasSize(2);
    ParseStatus status = result.status();
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_OK);
    assertThat(status.getParagraphs()).isZero();
    assertThat(status.getTables()).isZero();
  }

  // --- spreadsheet cell typing and formatting -----------------------------

  @Test
  void dateCellsArriveTypedWithExactInstant() throws Exception {
    Date when = Date.from(Instant.parse("2024-03-15T10:30:00Z"));
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var style = workbook.createCellStyle();
      style.setDataFormat(
          workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
      var cell = workbook.createSheet("Dates").createRow(0).createCell(0);
      cell.setCellValue(when);
      cell.setCellStyle(style);
      workbook.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "dates");
    SheetCell cell = result.eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet().getRows(0).getCells(0);
    assertThat(cell.hasDate()).as("date-formatted numeric arrives as a timestamp").isTrue();
    assertThat(cell.getDate().getSeconds())
        .as("Excel serial to epoch conversion must round-trip the authored instant")
        .isEqualTo(when.getTime() / 1000);
    assertThat(cell.getDate().getNanos()).isZero();
  }

  @Test
  void formulaErrorsSurfaceAsCanonicalErrorLiterals() throws Exception {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var row = workbook.createSheet("Errors").createRow(0);
      row.createCell(0).setCellFormula("1/0");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "errors");
    SheetCell cell = result.eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet().getRows(0).getCells(0);
    assertThat(cell.getFormula()).isEqualTo("1/0");
    assertThat(cell.getText())
        .as("cached error result maps to Excel's canonical literal")
        .isEqualTo("#DIV/0!");
  }

  @Test
  void numberFormattingIsAppliedToTheFormattedFieldOnly() throws Exception {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var style = workbook.createCellStyle();
      style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("0.00"));
      var cell = workbook.createSheet("Fmt").createRow(0).createCell(0);
      cell.setCellValue(3.14159);
      cell.setCellStyle(style);
      workbook.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "fmt");
    SheetCell cell = result.eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet().getRows(0).getCells(0);
    assertThat(cell.getNumber()).as("raw value keeps full precision").isEqualTo(3.14159);
    assertThat(cell.getFormatted())
        .as("DataFormatter output is wire-visible; POI must render 0.00 identically")
        .isEqualTo("3.14");
  }

  @Test
  void legacyXlsSharesTheTypedCellPathIncludingFormulas() throws Exception {
    byte[] bytes;
    try (HSSFWorkbook workbook = new HSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      workbook.createInformationProperties();
      workbook.getSummaryInformation().setTitle("Ledger");
      workbook.getSummaryInformation().setAuthor("Clerk");
      var row = workbook.createSheet("Legacy").createRow(0);
      row.createCell(0).setCellValue(10.0);
      row.createCell(1).setCellFormula("A1*3");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "xls-typed");
    var info = result.events().get(0).getDocumentInfo();
    assertThat(info.getFormat()).isEqualTo(DocumentFormat.DOCUMENT_FORMAT_LEGACY_XLS);
    assertThat(info.getMetadata().getTitle())
        .as("HPSF summary title maps to the same first-class field OOXML uses")
        .isEqualTo("Ledger");
    assertThat(info.getMetadata().getAuthor()).isEqualTo("Clerk");
    var cells = result.eventsOf(ParseEvent::hasSheet).get(0).getSheet().getRows(0);
    assertThat(cells.getCells(0).getNumber()).isEqualTo(10.0);
    assertThat(cells.getCells(1).getFormula()).isEqualTo("A1*3");
    assertThat(cells.getCells(1).getNumber())
        .as("HSSF cached formula result rides the typed value like XSSF")
        .isEqualTo(30.0);
  }

  // --- metadata mapping ---------------------------------------------------

  @Test
  void ooxmlPropertyTailKeepsDeclaredKeysAndTypes() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var core = document.getProperties().getCoreProperties();
      core.setSubjectProperty("Annual filing");
      core.setKeywords("finance, audit");
      core.setCategory("reports");
      var custom = document.getProperties().getCustomProperties();
      custom.addProperty("revision_number", 7);
      custom.addProperty("confidence", 0.75);
      custom.addProperty("codename", "osprey");
      document.createParagraph().createRun().setText("body");
      document.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "meta");
    var tail = result.events().get(0).getDocumentInfo().getMetadata().getTailList();
    assertThat(tail).anyMatch(entry -> entry.getKey().equals("subject")
        && entry.getValues(0).getStringValue().equals("Annual filing"));
    assertThat(tail).anyMatch(entry -> entry.getKey().equals("keywords")
        && entry.getValues(0).getStringValue().equals("finance, audit"));
    assertThat(tail).anyMatch(entry -> entry.getKey().equals("category")
        && entry.getValues(0).getStringValue().equals("reports"));
    assertThat(tail).as("custom int stays an int, never a string")
        .anyMatch(entry -> entry.getKey().equals("custom:revision_number")
            && entry.getValues(0).getIntValue() == 7);
    assertThat(tail).anyMatch(entry -> entry.getKey().equals("custom:confidence")
        && entry.getValues(0).getDoubleValue() == 0.75);
    assertThat(tail).anyMatch(entry -> entry.getKey().equals("custom:codename")
        && entry.getValues(0).getStringValue().equals("osprey"));
  }

  // --- slide placeholder routing ------------------------------------------

  @Test
  void pptxTitlePlaceholderRoutesToTitleNotTexts() throws Exception {
    byte[] bytes;
    try (XMLSlideShow show = new XMLSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      XSLFSlideLayout layout =
          show.getSlideMasters().get(0).getLayout(SlideLayout.TITLE_AND_CONTENT);
      assertThat(layout).as("default template must carry TITLE_AND_CONTENT").isNotNull();
      XSLFSlide slide = show.createSlide(layout);
      slide.getPlaceholder(0).setText("Deck Title");
      slide.getPlaceholder(1).setText("Body point");
      XSLFNotes notes = show.getNotesSlide(slide);
      for (XSLFTextShape shape : notes.getPlaceholders()) {
        if (shape.getTextType() == Placeholder.BODY) shape.setText("Speaker reminder");
      }
      show.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "pptx-title");
    var slide = result.eventsOf(ParseEvent::hasSlide).get(0).getSlide();
    assertThat(slide.getTitle()).isEqualTo("Deck Title");
    assertThat(slide.getTextsList())
        .as("the title must not be duplicated into texts")
        .containsExactly("Body point");
    assertThat(slide.getNotesList()).contains("Speaker reminder");
  }

  @Test
  void legacyPptTitleWithoutPlaceholderAtomStaysInTexts() throws Exception {
    byte[] bytes;
    try (HSLFSlideShow show = new HSLFSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var slide = show.createSlide();
      slide.addTitle().setText("Old Deck");
      HSLFTextBox box = slide.createTextBox();
      box.setText("legacy body");
      show.write(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parseOk(bytes, "ppt-title");
    var slide = result.eventsOf(ParseEvent::hasSlide).get(0).getSlide();
    assertThat(slide.getTitle()).isEqualTo("Old Deck");
    // HSLF reports no TITLE placeholder for a title shape that lacks an
    // OEPlaceholderAtom (POI-authored decks), so the title text also rides
    // texts. That duplication is the standing wire contract for such decks;
    // if a POI upgrade starts reporting the placeholder, texts would silently
    // lose an element, which this pin turns into a loud failure.
    assertThat(slide.getTextsList()).containsExactly("Old Deck", "legacy body");
  }

  // --- format detection and error mapping ---------------------------------

  @Test
  void truncatedOoxmlContainerIsInvalidArgument() throws Exception {
    byte[] whole;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText("will be cut");
      document.write(out);
      whole = out.toByteArray();
    }
    byte[] truncated = Arrays.copyOf(whole, whole.length / 2);
    ParseResult result = harness.parse(truncated, "cut", truncated.length);
    assertThat(result.error()).isNotNull();
    assertThat(((StatusRuntimeException) result.error()).getStatus().getCode())
        .as("zip magic with an unreadable container is a bad document, not an unknown format")
        .isEqualTo(Status.Code.INVALID_ARGUMENT);
  }

  @Test
  void ole2WithoutOfficeStreamIsUnimplemented() throws Exception {
    byte[] bytes;
    try (POIFSFileSystem container = new POIFSFileSystem();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      container.createDocument(
          new ByteArrayInputStream("not an office payload".getBytes()), "CustomStream");
      container.writeFilesystem(out);
      bytes = out.toByteArray();
    }
    ParseResult result = harness.parse(bytes, "ole2-alien", bytes.length);
    assertThat(result.error()).isNotNull();
    StatusRuntimeException status = (StatusRuntimeException) result.error();
    assertThat(status.getStatus().getCode()).isEqualTo(Status.Code.UNIMPLEMENTED);
    assertThat(status.getStatus().getDescription())
        .isEqualTo("OLE2 container without a known office stream");
  }

  @Test
  void serviceInfoPoiVersionTracksTheRunningLibrary() {
    var info = harness.blockingStub().getServiceInfo(GetServiceInfoRequest.getDefaultInstance());
    assertThat(info.getPoiVersion())
        .matches("\\d+\\.\\d+\\.\\d+")
        .isEqualTo(org.apache.poi.Version.getVersion());
  }
}
