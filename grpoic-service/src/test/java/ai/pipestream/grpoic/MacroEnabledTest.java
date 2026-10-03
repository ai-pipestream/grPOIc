package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.ParseEvent;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Macro-enabled packages (.docm, .xlsm, .pptm) and templates parse like
 * their plain siblings. The VBA project rides along as a package part that
 * nothing reads: it is neither run nor reported as an embedded object.
 */
class MacroEnabledTest {

  private static final String VBA_RELATIONSHIP =
      "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
  private static final byte[] VBA_PAYLOAD =
      "Attribute VB_Name = \"AutoOpen\"\nSub AutoOpen(): Shell \"calc\": End Sub"
          .getBytes(StandardCharsets.US_ASCII);

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  /**
   * Rewrites the main part's content type and adds a vbaProject part, which
   * is what turns a package into its macro-enabled (or template) form.
   */
  private static byte[] withMainType(POIXMLDocument document, String fromType, String toType,
                                     String vbaPart) throws Exception {
    byte[] plain;
    try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.write(out);
      plain = out.toByteArray();
    }
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(plain));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      container.replaceContentType(fromType, toType);
      if (vbaPart != null) {
        PackagePart vba = container.createPart(
            PackagingURIHelper.createPartName(vbaPart), "application/vnd.ms-office.vbaProject");
        try (OutputStream vbaOut = vba.getOutputStream()) {
          vbaOut.write(VBA_PAYLOAD);
        }
        PackagePart main = container.getPart(container.getRelationshipsByType(
            PackageRelationshipTypes.CORE_DOCUMENT).getRelationship(0));
        main.addRelationship(vba.getPartName(), TargetMode.INTERNAL, VBA_RELATIONSHIP);
      }
      container.save(out);
      return out.toByteArray();
    }
  }

  private static void assertNoVbaSurfaced(ParseResult result) {
    assertThat(result.eventsOf(ParseEvent::hasEmbeddedObject))
        .as("the VBA project is not an embedded object and is never read")
        .isEmpty();
    assertThat(result.events().toString()).doesNotContain("AutoOpen");
  }

  @Test
  void docmParsesLikeDocx() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument()) {
      document.createParagraph().createRun().setText("macro document body");
      bytes = withMainType(document, XWPFRelation.DOCUMENT.getContentType(),
          XWPFRelation.MACRO_DOCUMENT.getContentType(), "/word/vbaProject.bin");
    }
    ParseResult result = harness.parseOk(bytes, "docm");
    assertThat(result.events().get(0).getDocumentInfo().getFormat())
        .isEqualTo(DocumentFormat.DOCUMENT_FORMAT_DOCX);
    assertThat(result.eventsOf(ParseEvent::hasParagraph))
        .extracting(event -> event.getParagraph().getText())
        .containsExactly("macro document body");
    assertNoVbaSurfaced(result);
  }

  @Test
  void xlsmParsesLikeXlsx() throws Exception {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Macros").createRow(0).createCell(0).setCellValue("macro cell");
      bytes = withMainType(workbook, XSSFRelation.WORKBOOK.getContentType(),
          XSSFRelation.MACROS_WORKBOOK.getContentType(), "/xl/vbaProject.bin");
    }
    ParseResult result = harness.parseOk(bytes, "xlsm");
    assertThat(result.events().get(0).getDocumentInfo().getFormat())
        .isEqualTo(DocumentFormat.DOCUMENT_FORMAT_XLSX);
    assertThat(result.eventsOf(ParseEvent::hasSheet).get(0).getSheet()
        .getRows(0).getCells(0).getText()).isEqualTo("macro cell");
    assertNoVbaSurfaced(result);
  }

  @Test
  void pptmParsesLikePptx() throws Exception {
    byte[] bytes;
    try (XMLSlideShow show = new XMLSlideShow()) {
      show.createSlide().createTextBox().setText("macro slide");
      bytes = withMainType(show, XSLFRelation.MAIN.getContentType(),
          XSLFRelation.PRESENTATION_MACRO.getContentType(), "/ppt/vbaProject.bin");
    }
    ParseResult result = harness.parseOk(bytes, "pptm");
    assertThat(result.events().get(0).getDocumentInfo().getFormat())
        .isEqualTo(DocumentFormat.DOCUMENT_FORMAT_PPTX);
    assertThat(result.eventsOf(ParseEvent::hasSlide).get(0).getSlide().getTextsList())
        .containsExactly("macro slide");
    assertNoVbaSurfaced(result);
  }

  @Test
  void macroEnabledTemplatesParseToo() throws Exception {
    byte[] bytes;
    try (XWPFDocument document = new XWPFDocument()) {
      document.createParagraph().createRun().setText("template body");
      bytes = withMainType(document, XWPFRelation.DOCUMENT.getContentType(),
          XWPFRelation.MACRO_TEMPLATE_DOCUMENT.getContentType(), null);
    }
    assertThat(harness.parseOk(bytes, "dotm").eventsOf(ParseEvent::hasParagraph))
        .extracting(event -> event.getParagraph().getText())
        .containsExactly("template body");
  }
}
