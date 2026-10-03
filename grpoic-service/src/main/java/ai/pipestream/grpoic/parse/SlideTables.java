package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.Table;
import ai.pipestream.poi.v1.TableCell;
import ai.pipestream.poi.v1.TableRow;
import org.apache.poi.sl.usermodel.TableShape;

/**
 * A slide's native table under the wire's span contract: each merged region
 * is one anchor cell carrying both spans, covered positions are not
 * repeated, and every other grid position holds a cell, so a consumer that
 * places cells by counting spans lands each one where the slide shows it.
 *
 * <p><b>Spans are trusted only as far as the grid allows.</b> A span is cut
 * at the table's edge and at the first position an earlier merge already
 * covers; a position flagged as merged that no anchor covers becomes an
 * empty cell. A run of such positions arrives as one empty cell spanning
 * it, so a crafted table cannot multiply the output.
 *
 * <p><b>Bounded grid.</b> The grid is cut at {@link #MAX_COLUMNS} columns
 * and {@link #MAX_ROWS} rows; a slide table never legitimately comes near
 * either.
 */
final class SlideTables {

  static final int MAX_COLUMNS = WordTables.MAX_SPAN;
  static final int MAX_ROWS = 10_000;

  private SlideTables() {}

  static Table convert(TableShape<?, ?> table) {
    int rows = Math.min(Math.max(table.getNumberOfRows(), 0), MAX_ROWS);
    int columns = Math.min(Math.max(table.getNumberOfColumns(), 0), MAX_COLUMNS);
    // The first row each column is free again, as a consumer tracks it.
    int[] occupiedUntil = new int[columns];
    Table.Builder converted = Table.newBuilder();
    for (int row = 0; row < rows; row++) {
      TableRow.Builder convertedRow = TableRow.newBuilder();
      int column = 0;
      while (column < columns) {
        if (occupiedUntil[column] > row) {
          column++;
          continue;
        }
        var cell = table.getCell(row, column);
        if (cell == null || cell.isMerged()) {
          int gap = 1;
          while (column + gap < columns && occupiedUntil[column + gap] <= row) {
            var next = table.getCell(row, column + gap);
            if (next != null && !next.isMerged()) break;
            gap++;
          }
          convertedRow.addCells(TableCell.newBuilder().setText("").setRowSpan(1).setColSpan(gap));
          column += gap;
          continue;
        }
        int colSpan = Math.min(Math.max(cell.getGridSpan(), 1), columns - column);
        for (int covered = 1; covered < colSpan; covered++) {
          if (occupiedUntil[column + covered] > row) {
            colSpan = covered;
            break;
          }
        }
        int rowSpan = Math.min(Math.max(cell.getRowSpan(), 1), rows - row);
        for (int held = column; held < column + colSpan; held++) {
          occupiedUntil[held] = row + rowSpan;
        }
        String text = cell.getText();
        convertedRow.addCells(
            TableCell.newBuilder()
                .setText(text == null ? "" : text)
                .setRowSpan(rowSpan)
                .setColSpan(colSpan));
        column += colSpan;
      }
      converted.addRows(convertedRow);
    }
    return converted.build();
  }
}
