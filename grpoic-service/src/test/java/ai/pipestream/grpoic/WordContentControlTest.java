package ai.pipestream.grpoic;

import static ai.pipestream.grpoic.OoxmlFixtures.docxWithBody;
import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * DOCX content controls (w:sdt) are opened up rather than dropped: a
 * table of contents or a template's fill-in region comes out as the
 * paragraphs and tables it holds, in document order, and an inline control
 * keeps its text inside the paragraph.
 */
class WordContentControlTest {

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static String paragraph(String style, String text) {
    return "<w:p>" + (style.isEmpty() ? "" : "<w:pPr><w:pStyle w:val=\"" + style + "\"/></w:pPr>")
        + "<w:r><w:t xml:space=\"preserve\">" + text + "</w:t></w:r></w:p>";
  }

  @Test
  void blockControlsEmitTheirParagraphsAndTablesInPlace() throws Exception {
    byte[] bytes = docxWithBody(
        paragraph("", "Before")
            + "<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/>"
            + "</w:docPartObj></w:sdtPr><w:sdtContent>"
            + paragraph("TOCHeading", "Contents")
            + "<w:sdt><w:sdtPr/><w:sdtContent>" + paragraph("TOC1", "Chapter one")
            + "</w:sdtContent></w:sdt>"
            + "<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>"
            + paragraph("", "in control") + "</w:tc></w:tr></w:tbl>"
            + "</w:sdtContent></w:sdt>"
            + paragraph("", "After"));

    ParseResult result = harness.parseOk(bytes, "sdt-block");
    var body = result.events().stream()
        .filter(event -> event.hasParagraph() || event.hasTable())
        .toList();
    assertThat(body).extracting(event -> event.hasTable()
            ? "table:" + event.getTable().getRows(0).getCells(0).getText()
            : event.getParagraph().getStyle() + ":" + event.getParagraph().getText())
        .as("control content arrives where the control sits, nested controls flattened")
        .containsExactly(
            ":Before", "TOCHeading:Contents", "TOC1:Chapter one", "table:in control", ":After");
    ParseStatus status = result.status();
    assertThat(status.getParagraphs()).isEqualTo(4);
    assertThat(status.getTables()).isEqualTo(1);
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_OK);
  }

  @Test
  void inlineControlTextStaysInItsParagraph() throws Exception {
    byte[] bytes = docxWithBody(
        "<w:p><w:r><w:t xml:space=\"preserve\">Customer: </w:t></w:r>"
            + "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>ACME Corp</w:t></w:r></w:sdtContent>"
            + "</w:sdt></w:p>");
    assertThat(harness.parseOk(bytes, "sdt-inline").eventsOf(ParseEvent::hasParagraph))
        .extracting(event -> event.getParagraph().getText())
        .containsExactly("Customer: ACME Corp");
  }
}
