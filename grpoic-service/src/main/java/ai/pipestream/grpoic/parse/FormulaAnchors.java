package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseStatus;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;

/**
 * The formulas the rows of one streamed worksheet refer back to: the
 * masters of shared formulas and the anchors of array formulas, kept only
 * while a later row can still need them.
 *
 * <p><b>Why not the sheet.</b> Bound to an {@code XSSFSheet}, every master
 * and every array range is recorded in the sheet's private state, which
 * grows with the sheet and is only dropped with the workbook; array ranges
 * are also a list every cell of the sheet scans, so a sheet of many array
 * formulas costs time quadratic in its size. Here a row's references are
 * taken out of the row before it is bound, so the sheet records nothing,
 * and kept in this class instead.
 *
 * <p><b>Why it stays small.</b> Rows are read in stored order, which is
 * ascending in every real writer. A group is dropped once the rows have
 * moved past its last row, so what is held is only the groups that cross
 * the current row. Groups do not overlap in a valid sheet, so that is at
 * most one per column; array anchors are keyed by their first column and
 * shared masters are capped at the column count, so even a hostile sheet
 * holds at most {@link #MAX_LIVE} of each. A row stored out of order after
 * its group was dropped renders the cell as POI would without the group:
 * an array-covered cell as its cached value, a shared-formula cell skipped
 * with a warning.
 *
 * <p><b>Same output.</b> An array-covered cell without a formula of its
 * own is given its anchor's formula text, which is what the usermodel
 * returns for it; a shared-formula master keeps its type and is answered
 * from here through {@link XlsxSheets.ScratchSheet#getSharedFormula(int)},
 * with its range adjusted exactly as {@code XSSFSheet} adjusts it.
 */
final class FormulaAnchors {

  /** Live groups of each kind held at once: one per column in a valid sheet. */
  static final int MAX_LIVE = SpreadsheetVersion.EXCEL2007.getMaxColumns();

  /**
   * A group of cells one formula covers, dropped once the rows pass
   * {@code lastRow}; {@code key} is the array's first column or the shared
   * formula's si.
   */
  private record Group(long order, int lastRow, boolean array, int key) {}

  private static final Comparator<Group> BY_LAST_ROW =
      Comparator.comparingInt(Group::lastRow).thenComparingLong(Group::order);

  private record ArrayFormula(CellRangeAddress range, String formula, Group group) {}

  private record SharedFormula(CTCellFormula master, Group group) {}

  private final String sheetName;
  private final ParseStatus.Builder status;
  /** Live array formulas by their first column. */
  private final TreeMap<Integer, ArrayFormula> arrays = new TreeMap<>();
  /** Live shared-formula masters by their si. */
  private final Map<Integer, SharedFormula> shared = new HashMap<>();
  private final TreeSet<Group> byLastRow = new TreeSet<>(BY_LAST_ROW);
  private long order;
  private boolean warnedFull;

  FormulaAnchors(String sheetName, ParseStatus.Builder status) {
    this.sheetName = sheetName;
    this.status = status;
  }

  /** The shared formula {@code si} refers to, or null when it is not held. */
  CTCellFormula shared(int si) {
    SharedFormula formula = shared.get(si);
    return formula == null ? null : formula.master();
  }

  /** Live groups of both kinds, for tests. */
  int live() {
    return arrays.size() + shared.size();
  }

  /**
   * Prepares one stored row (zero-based {@code rowIndex}) for binding:
   * drops the groups the rows have passed, takes this row's masters and
   * anchors out of it, and gives array-covered cells their formula. Cells
   * are located as {@code XSSFRow} locates them; a row whose references POI
   * cannot read is left as it is, so binding reports it as before.
   */
  void prepare(CTRow row, int rowIndex) {
    retire(rowIndex);
    CTCell[] cells = row.getCArray();
    int[] columns = new int[cells.length];
    int last = -1;
    for (int index = 0; index < cells.length; index++) {
      CTCell cell = cells[index];
      try {
        columns[index] = cell.getR() != null ? new CellReference(cell.getR()).getCol() : last + 1;
      } catch (RuntimeException unreadable) {
        return;
      }
      last = Math.max(last, columns[index]);
    }
    for (int index = 0; index < cells.length; index++) {
      if (cells[index].isSetF()) open(cells[index].getF(), rowIndex, columns[index]);
    }
    if (arrays.isEmpty()) return;
    for (int index = 0; index < cells.length; index++) {
      cover(cells[index], rowIndex, columns[index]);
    }
  }

  private void retire(int rowIndex) {
    while (!byLastRow.isEmpty() && byLastRow.first().lastRow() < rowIndex) {
      Group passed = byLastRow.pollFirst();
      if (passed.array()) {
        arrays.remove(passed.key());
      } else {
        shared.remove(passed.key());
      }
    }
  }

  /** Records a master or an anchor and strips what would make the sheet record it. */
  private void open(CTCellFormula formula, int rowIndex, int column) {
    if (formula.getT() == STCellFormulaType.ARRAY && formula.isSetRef()) {
      openArray(formula, rowIndex, column);
    } else if (formula.getT() == STCellFormulaType.SHARED && formula.isSetRef()
        && formula.getStringValue() != null) {
      openShared(formula, rowIndex, column);
    }
  }

  private void openArray(CTCellFormula formula, int rowIndex, int column) {
    String where = "sheet '" + sheetName + "' array formula at "
        + new CellReference(rowIndex, column).formatAsString();
    CellRangeAddress range;
    try {
      range = CellRangeAddress.valueOf(formula.getRef());
    } catch (RuntimeException unreadable) {
      DocumentFaults.skip(status, where, unreadable);
      range = null;
    }
    // The anchor itself renders its own text whether or not it is an array.
    formula.unsetRef();
    formula.unsetT();
    if (range == null || range.getLastRow() < rowIndex) return;
    Group group = new Group(order++, range.getLastRow(), true, range.getFirstColumn());
    ArrayFormula replaced = arrays.put(range.getFirstColumn(),
        new ArrayFormula(range, formula.getStringValue(), group));
    if (replaced != null) byLastRow.remove(replaced.group());
    byLastRow.add(group);
  }

  private void openShared(CTCellFormula formula, int rowIndex, int column) {
    CTCellFormula master = (CTCellFormula) formula.copy();
    // Without a ref the sheet does not record the master; the copy keeps it.
    formula.unsetRef();
    String where = "sheet '" + sheetName + "' shared formula at "
        + new CellReference(rowIndex, column).formatAsString();
    if (master.getSi() > Integer.MAX_VALUE) {
      // si is unsigned 32-bit; POI keys shared formulas by int and cannot hold it.
      DocumentFaults.warn(status, where + " skipped: its si is out of range");
      return;
    }
    int si = (int) master.getSi();
    CellRangeAddress range;
    try {
      range = CellRangeAddress.valueOf(master.getRef());
    } catch (RuntimeException unreadable) {
      DocumentFaults.skip(status, where, unreadable);
      return;
    }
    // XSSFSheet's adjustment for a master that is not the range's top left.
    if (column > range.getFirstColumn() || rowIndex > range.getFirstRow()) {
      range = new CellRangeAddress(
          Math.max(rowIndex, range.getFirstRow()), Math.max(rowIndex, range.getLastRow()),
          Math.max(column, range.getFirstColumn()), Math.max(column, range.getLastColumn()));
      master.setRef(range.formatAsString());
    }
    SharedFormula replaced = shared.get(si);
    if (replaced == null && shared.size() >= MAX_LIVE) {
      if (!warnedFull) {
        warnedFull = true;
        DocumentFaults.warn(status, "sheet '" + sheetName + "' has more than " + MAX_LIVE
            + " shared formulas open at once; cells using the rest were skipped");
      }
      return;
    }
    Group group = new Group(order++, range.getLastRow(), false, si);
    shared.put(si, new SharedFormula(master, group));
    if (replaced != null) byLastRow.remove(replaced.group());
    byLastRow.add(group);
  }

  /**
   * Gives a cell inside a live array range its anchor's formula when it has
   * none of its own, as the usermodel does when asked for its formula.
   */
  private void cover(CTCell cell, int rowIndex, int column) {
    if (cell.isSetF() && !cell.getF().getStringValue().isEmpty()) return;
    Map.Entry<Integer, ArrayFormula> candidate = arrays.floorEntry(column);
    if (candidate == null || !candidate.getValue().range().isInRange(rowIndex, column)) return;
    CTCellFormula formula = cell.isSetF() ? cell.getF() : cell.addNewF();
    if (formula.isSetT()) formula.unsetT();
    if (formula.isSetRef()) formula.unsetRef();
    if (formula.isSetSi()) formula.unsetSi();
    formula.setStringValue(candidate.getValue().formula());
  }
}
