package ai.pipestream.grpoic;

import ai.pipestream.poi.v1.Table;
import ai.pipestream.poi.v1.TableCell;
import java.util.ArrayList;
import java.util.List;

/**
 * Places a wire Table on its grid exactly the way gRParse's client-side fold
 * does (poi_collector.cpp): a row's cells go left to right, each at the next
 * column that no row span from an earlier row still holds. Asserting on the
 * placed positions proves the wire contract holds for that consumer: no
 * column shifts, no covered position repeated.
 */
final class TableGrid {

  record Placed(String text, int row, int column, int rowSpan, int colSpan) {}

  private TableGrid() {}

  static List<Placed> place(Table table) {
    List<Placed> placed = new ArrayList<>();
    List<Integer> occupiedUntil = new ArrayList<>();
    for (int row = 0; row < table.getRowsCount(); row++) {
      int column = 0;
      for (TableCell cell : table.getRows(row).getCellsList()) {
        while (column < occupiedUntil.size() && occupiedUntil.get(column) > row) column++;
        int rowSpan = Math.max(1, cell.getRowSpan());
        int colSpan = Math.max(1, cell.getColSpan());
        while (occupiedUntil.size() < column + colSpan) occupiedUntil.add(0);
        for (int held = column; held < column + colSpan; held++) {
          occupiedUntil.set(held, row + rowSpan);
        }
        placed.add(new Placed(cell.getText(), row, column, rowSpan, colSpan));
        column += colSpan;
      }
    }
    return placed;
  }
}
