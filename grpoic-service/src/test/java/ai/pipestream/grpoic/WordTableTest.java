package ai.pipestream.grpoic;

import static ai.pipestream.grpoic.OoxmlFixtures.docxWithBody;
import static ai.pipestream.grpoic.OoxmlFixtures.tc;
import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.TableGrid.Placed;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.Table;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * DOCX tables against the wire's span contract, placed the way gRParse's
 * fold places them: vertical merges become one anchor with a row span and
 * the covered positions disappear, row padding and content-control cells
 * hold their columns, and spans stay within bounds.
 */
class WordTableTest {

  private static final String GRID3 =
      "<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>";

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static Table onlyTable(String tableXml) throws Exception {
    var tables = harness.parseOk(docxWithBody(tableXml), "table").eventsOf(ParseEvent::hasTable);
    assertThat(tables).hasSize(1);
    return tables.get(0).getTable();
  }

  @Test
  void verticalMergeSpansRowsAndRowPaddingHoldsItsColumns() throws Exception {
    Table table = onlyTable("<w:tbl>" + GRID3
        + "<w:tr>" + tc("<w:vMerge w:val=\"restart\"/>", "North") + tc("Q1") + tc("10") + "</w:tr>"
        + "<w:tr>" + tc("<w:vMerge/>", "") + tc("Q2") + tc("20") + "</w:tr>"
        + "<w:tr>" + tc("<w:vMerge w:val=\"continue\"/>", "") + tc("Q3") + tc("30") + "</w:tr>"
        + "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>" + tc("Total") + tc("60") + "</w:tr>"
        + "</w:tbl>");

    assertThat(table.getRows(0).getCells(0).getRowSpan())
        .as("the restart cell anchors all three merged rows")
        .isEqualTo(3);
    assertThat(table.getRows(1).getCellsCount())
        .as("covered positions are not repeated")
        .isEqualTo(2);
    assertThat(TableGrid.place(table)).containsExactly(
        new Placed("North", 0, 0, 3, 1),
        new Placed("Q1", 0, 1, 1, 1),
        new Placed("10", 0, 2, 1, 1),
        new Placed("Q2", 1, 1, 1, 1),
        new Placed("20", 1, 2, 1, 1),
        new Placed("Q3", 2, 1, 1, 1),
        new Placed("30", 2, 2, 1, 1),
        new Placed("", 3, 0, 1, 1),
        new Placed("Total", 3, 1, 1, 1),
        new Placed("60", 3, 2, 1, 1));
  }

  @Test
  void mergeInAMiddleColumnAndTrailingPaddingKeepPositions() throws Exception {
    Table table = onlyTable("<w:tbl>" + GRID3
        + "<w:tr>" + tc("a") + tc("<w:vMerge w:val=\"restart\"/><w:gridSpan w:val=\"2\"/>", "wide")
        + "</w:tr>"
        + "<w:tr>" + tc("b") + tc("<w:vMerge/><w:gridSpan w:val=\"2\"/>", "") + "</w:tr>"
        + "<w:tr><w:trPr><w:gridAfter w:val=\"2\"/></w:trPr>" + tc("c") + "</w:tr>"
        + "</w:tbl>");
    assertThat(TableGrid.place(table)).containsExactly(
        new Placed("a", 0, 0, 1, 1),
        new Placed("wide", 0, 1, 2, 2),
        new Placed("b", 1, 0, 1, 1),
        new Placed("c", 2, 0, 1, 1),
        new Placed("", 2, 1, 1, 2));
  }

  @Test
  void contentControlCellsAndRowsAreKeptInPlace() throws Exception {
    Table table = onlyTable("<w:tbl>" + GRID3
        + "<w:tr>" + tc("A1")
        + "<w:sdt><w:sdtPr/><w:sdtContent>" + tc("B1 wrapped") + "</w:sdtContent></w:sdt>"
        + tc("C1") + "</w:tr>"
        + "<w:sdt><w:sdtPr/><w:sdtContent>"
        + "<w:tr>" + tc("A2") + tc("B2") + tc("C2") + "</w:tr>"
        + "</w:sdtContent></w:sdt>"
        + "<w:tr><w:tc><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>in a block control</w:t>"
        + "</w:r></w:p></w:sdtContent></w:sdt></w:tc>" + tc("B3") + tc("C3") + "</w:tr>"
        + "</w:tbl>");
    assertThat(TableGrid.place(table)).containsExactly(
        new Placed("A1", 0, 0, 1, 1),
        new Placed("B1 wrapped", 0, 1, 1, 1),
        new Placed("C1", 0, 2, 1, 1),
        new Placed("A2", 1, 0, 1, 1),
        new Placed("B2", 1, 1, 1, 1),
        new Placed("C2", 1, 2, 1, 1),
        new Placed("in a block control", 2, 0, 1, 1),
        new Placed("B3", 2, 1, 1, 1),
        new Placed("C3", 2, 2, 1, 1));
  }

  @Test
  void continuationThatCannotJoinStaysAnOrdinaryCell() throws Exception {
    Table table = onlyTable("<w:tbl>" + GRID3
        + "<w:tr>" + tc("<w:gridSpan w:val=\"2\"/>", "two wide") + tc("x") + "</w:tr>"
        + "<w:tr>" + tc("<w:vMerge/>", "narrow") + tc("y") + tc("z") + "</w:tr>"
        + "</w:tbl>");
    assertThat(table.getRows(0).getCells(0).getRowSpan())
        .as("a continuation narrower than the cell above does not merge into it")
        .isEqualTo(1);
    assertThat(TableGrid.place(table)).containsExactly(
        new Placed("two wide", 0, 0, 1, 2),
        new Placed("x", 0, 2, 1, 1),
        new Placed("narrow", 1, 0, 1, 1),
        new Placed("y", 1, 1, 1, 1),
        new Placed("z", 1, 2, 1, 1));
  }

  @Test
  void spansAndPaddingAreClampedToTheSaneMaximum() throws Exception {
    Table table = onlyTable("<w:tbl>" + GRID3
        + "<w:tr>" + tc("<w:gridSpan w:val=\"4294967295\"/>", "huge") + "</w:tr>"
        + "<w:tr>" + tc("<w:gridSpan w:val=\"-3\"/>", "negative") + tc("<w:gridSpan/>", "no value")
        + "</w:tr>"
        + "<w:tr><w:trPr><w:gridBefore w:val=\"99999999999\"/></w:trPr>" + tc("late") + "</w:tr>"
        + "</w:tbl>");
    assertThat(table.getRows(0).getCells(0).getColSpan()).isEqualTo(1024);
    assertThat(table.getRows(1).getCells(0).getColSpan()).isEqualTo(1);
    assertThat(table.getRows(1).getCells(1).getColSpan()).isEqualTo(1);
    assertThat(table.getRows(2).getCells(0).getColSpan())
        .as("row padding is clamped too, and arrives as a single cell")
        .isEqualTo(1024);
    assertThat(table.getRows(2).getCellsCount()).isEqualTo(2);
  }
}
