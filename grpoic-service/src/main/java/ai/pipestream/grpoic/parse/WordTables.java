package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.Table;
import ai.pipestream.poi.v1.TableCell;
import ai.pipestream.poi.v1.TableRow;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

/**
 * Lays a DOCX table onto its column grid and emits it under the wire's span
 * contract: the anchor of a merged region carries the text and both spans,
 * the positions it covers are not repeated.
 *
 * <p><b>Vertical merges.</b> Word marks the top cell of a vertical merge
 * {@code w:vMerge="restart"} and every cell below it {@code w:vMerge}
 * (continue). A continuation joins the cell directly above it when that cell
 * starts at the same grid column with the same width, and is then dropped;
 * the anchor's row span grows by one. A continuation with nothing to join
 * stays an ordinary cell, so a malformed merge can never move a column.
 *
 * <p><b>Rows that start late or end early.</b> {@code w:gridBefore} and
 * {@code w:gridAfter} leave grid columns empty at a row's edges. A consumer
 * places cells by counting spans from the left, so the leading gap must be
 * held by something: each gap arrives as one empty cell spanning it. One
 * cell rather than one per column, so a crafted padding cannot multiply the
 * output.
 *
 * <p><b>Spans are clamped.</b> {@code w:gridSpan} and the row padding are
 * attacker-controlled integers; each is clamped to {@code [1, MAX_SPAN]}
 * (padding to {@code [0, MAX_SPAN]}) before it reaches a {@code uint32}.
 * Word itself caps a table at 63 columns.
 */
final class WordTables {

  /** The widest span or row padding a cell may claim. */
  static final int MAX_SPAN = 1024;

  private WordTables() {}

  static Table convert(XWPFTable table) {
    IBody body = table.getBody();
    List<List<Slot>> rows = new ArrayList<>();
    // The cell of the previous row that starts at each grid column; the
    // candidates a continuation in this row can join.
    Map<Integer, Slot> above = Map.of();
    for (CTRow row : rows(table)) {
      List<Slot> cells = new ArrayList<>();
      Map<Integer, Slot> starts = new HashMap<>();
      int before = padding(row.getTrPr(), true);
      if (before > 0) cells.add(new Slot("", before));
      int column = before;
      for (CTTc cell : cells(row)) {
        int span = gridSpan(cell.getTcPr());
        Slot anchor = above.get(column);
        if (continuesMerge(cell.getTcPr()) && anchor != null && anchor.colSpan == span) {
          anchor.rowSpan++;
          starts.put(column, anchor);
        } else {
          Slot slot = new Slot(text(cell, body), span);
          cells.add(slot);
          starts.put(column, slot);
        }
        column += span;
      }
      int after = padding(row.getTrPr(), false);
      if (after > 0) cells.add(new Slot("", after));
      rows.add(cells);
      above = starts;
    }

    Table.Builder converted = Table.newBuilder();
    for (List<Slot> cells : rows) {
      TableRow.Builder convertedRow = TableRow.newBuilder();
      for (Slot slot : cells) {
        convertedRow.addCells(
            TableCell.newBuilder()
                .setText(slot.text)
                .setRowSpan(slot.rowSpan)
                .setColSpan(slot.colSpan));
      }
      converted.addRows(convertedRow);
    }
    return converted.build();
  }

  /** The table's rows in order, including rows wrapped in content controls. */
  private static List<CTRow> rows(XWPFTable table) {
    List<CTRow> rows = new ArrayList<>();
    ContentControls.forEachChild(table.getCTTbl(), child -> {
      if (child instanceof CTRow row) rows.add(row);
    });
    return rows;
  }

  /** The row's cells in order, including cells wrapped in content controls. */
  private static List<CTTc> cells(CTRow row) {
    List<CTTc> cells = new ArrayList<>();
    ContentControls.forEachChild(row, child -> {
      if (child instanceof CTTc cell) cells.add(cell);
    });
    return cells;
  }

  /**
   * The cell's paragraphs concatenated, as {@code XWPFTableCell.getText()}
   * does, with paragraphs inside content controls counted as the cell's own.
   * Nested tables stay out of the text, as they always have.
   */
  private static String text(CTTc cell, IBody body) {
    StringBuilder text = new StringBuilder();
    ContentControls.forEachChild(cell, child -> {
      if (child instanceof CTP paragraph) text.append(new XWPFParagraph(paragraph, body).getText());
    });
    return text.toString();
  }

  private static int gridSpan(CTTcPr properties) {
    if (properties == null || !properties.isSetGridSpan()) return 1;
    return clamp(properties.getGridSpan().getVal(), 1);
  }

  private static boolean continuesMerge(CTTcPr properties) {
    if (properties == null || !properties.isSetVMerge()) return false;
    // An absent val means continue.
    return properties.getVMerge().getVal() != STMerge.RESTART;
  }

  private static int padding(CTTrPr properties, boolean before) {
    if (properties == null) return 0;
    int count = before ? properties.sizeOfGridBeforeArray() : properties.sizeOfGridAfterArray();
    if (count == 0) return 0;
    CTDecimalNumber value =
        before ? properties.getGridBeforeArray(0) : properties.getGridAfterArray(0);
    return clamp(value.getVal(), 0);
  }

  private static int clamp(BigInteger value, int min) {
    if (value == null || value.compareTo(BigInteger.valueOf(min)) < 0) return min;
    if (value.compareTo(BigInteger.valueOf(MAX_SPAN)) > 0) return MAX_SPAN;
    return value.intValue();
  }

  /** One emitted cell; the row span grows as continuations join it. */
  private static final class Slot {
    final String text;
    final int colSpan;
    int rowSpan = 1;

    Slot(String text, int colSpan) {
      this.text = text;
      this.colSpan = colSpan;
    }
  }
}
