package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.grpoic.TableGrid.Placed;
import ai.pipestream.poi.v1.ParseEvent;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayOutputStream;
import org.apache.poi.hslf.usermodel.HSLFGroupShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Slide content that sits below the top level: text inside group shapes,
 * and native tables, which follow their slide as Table events carrying the
 * slide's index and obeying the same span contract as DOCX tables.
 */
class SlideShapesTest {

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  @Test
  void pptxGroupedTextAndMergedTablesAreKept() throws Exception {
    byte[] bytes;
    try (XMLSlideShow show = new XMLSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      show.createSlide();
      XSLFSlide slide = show.createSlide();
      XSLFTextBox loose = slide.createTextBox();
      loose.setAnchor(new Rectangle2D.Double(10, 10, 300, 40));
      loose.setText("loose");
      XSLFGroupShape group = slide.createGroup();
      group.createTextBox().setText("grouped bullet");
      group.createGroup().createTextBox().setText("nested deeper");
      XSLFTable table = slide.createTable(3, 3);
      String[][] texts = {{"Region", "Q1", "Q2"}, {"", "10", "20"}, {"Total", "30", ""}};
      for (int row = 0; row < 3; row++) {
        for (int column = 0; column < 3; column++) {
          table.getCell(row, column).setText(texts[row][column]);
        }
      }
      table.mergeCells(0, 1, 0, 0);
      table.mergeCells(2, 2, 1, 2);
      show.write(out);
      bytes = out.toByteArray();
    }

    ParseResult result = harness.parseOk(bytes, "pptx-shapes");
    var slides = result.eventsOf(ParseEvent::hasSlide);
    assertThat(slides).hasSize(2);
    assertThat(slides.get(1).getSlide().getTextsList())
        .as("group members count as the slide's text, in shape order")
        .containsExactly("loose", "grouped bullet", "nested deeper");

    var tables = result.eventsOf(ParseEvent::hasTable);
    assertThat(tables).hasSize(1);
    var table = tables.get(0).getTable();
    assertThat(table.hasSlideIndex()).isTrue();
    assertThat(table.getSlideIndex()).isEqualTo(1);
    int slideEvent = result.events().indexOf(slides.get(1));
    assertThat(result.events().indexOf(tables.get(0)))
        .as("a slide's tables follow its Slide event")
        .isEqualTo(slideEvent + 1);
    assertThat(TableGrid.place(table)).containsExactly(
        new Placed("Region", 0, 0, 2, 1),
        new Placed("Q1", 0, 1, 1, 1),
        new Placed("Q2", 0, 2, 1, 1),
        new Placed("10", 1, 1, 1, 1),
        new Placed("20", 1, 2, 1, 1),
        new Placed("Total", 2, 0, 1, 1),
        new Placed("30", 2, 1, 1, 2));
    assertThat(result.status().getTables()).isEqualTo(1);
    assertThat(result.status().getSlides()).isEqualTo(2);
  }

  @Test
  void legacyPptGroupedTextAndTablesAreKept() throws Exception {
    byte[] bytes;
    try (HSLFSlideShow show = new HSLFSlideShow();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      HSLFSlide slide = show.createSlide();
      HSLFGroupShape group = slide.createGroup();
      group.createTextBox().setText("legacy grouped");
      HSLFTable table = slide.createTable(2, 2);
      table.getCell(0, 0).setText("a");
      table.getCell(0, 1).setText("b");
      table.getCell(1, 0).setText("c");
      table.getCell(1, 1).setText("d");
      show.write(out);
      bytes = out.toByteArray();
    }

    ParseResult result = harness.parseOk(bytes, "ppt-shapes");
    assertThat(result.eventsOf(ParseEvent::hasSlide).get(0).getSlide().getTextsList())
        .as("table cells are not mistaken for loose text boxes")
        .containsExactly("legacy grouped");
    var table = result.eventsOf(ParseEvent::hasTable).get(0).getTable();
    assertThat(table.getSlideIndex()).isZero();
    assertThat(TableGrid.place(table)).extracting(Placed::text).containsExactly("a", "b", "c", "d");
  }
}
