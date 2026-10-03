package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.CellRange;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetCell;
import ai.pipestream.poi.v1.SheetRow;
import com.google.protobuf.CodedOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Supplier;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Spreadsheets through the common SS interface, so XSSF (.xlsx) and HSSF
 * (.xls) share one cell path. XLSX sheets stream a row at a time
 * ({@link XlsxSheets}); XLS, whose format caps a sheet at 65,536 rows, goes
 * through HSSF's usermodel. Formula cells are never evaluated; they carry
 * the formula source and the workbook's cached result.
 */
final class SpreadsheetParser {

  /**
   * Serialized row bytes per Sheet event when the client asked for sheet
   * batches. A batch closes at the first row that reaches it, so a message
   * is this plus at most one row.
   */
  static final int BATCH_BYTES = 1 << 20;
  /**
   * Serialized row bytes one Sheet event may hold when the client did not
   * ask for batches. The text budget alone would let a sheet of short cells
   * grow to tens of millions of cells in one builder; past this the parse
   * fails with {@link DocumentTooLargeException} and a hint to set
   * sheet_batches.
   */
  static final long UNBATCHED_SHEET_BYTES = 256L << 20;
  /** Merged regions kept per sheet; real sheets stay far below. */
  static final int MAX_MERGED_REGIONS = 100_000;
  /**
   * Characters every emitted cell costs on top of its text, about its
   * smallest size on the wire. Without it a cell with no text (an empty
   * shared string) would be free, and a sheet could emit millions of them.
   */
  static final int CELL_COST = 4;

  private SpreadsheetParser() {}

  /** A workbook loaded whole by its usermodel (XLS). */
  static void parse(Workbook workbook, ParseOptions options, Consumer<ParseEvent> emit,
                    ParseStatus.Builder status) {
    parse(workbook, options, emit, status, UNBATCHED_SHEET_BYTES);
  }

  /** As above with the unbatched event cap given, so tests can reach it with a small sheet. */
  static void parse(Workbook workbook, ParseOptions options, Consumer<ParseEvent> emit,
                    ParseStatus.Builder status, long unbatchedSheetBytes) {
    Conversion conversion = new Conversion(options, emit, status, unbatchedSheetBytes);
    for (int index = 0; index < workbook.getNumberOfSheets(); index++) {
      Sheet sheet = workbook.getSheetAt(index);
      boolean hidden = workbook.isSheetHidden(index) || workbook.isSheetVeryHidden(index);
      conversion.sheet(index, sheet.getSheetName(), hidden, sheet.iterator(),
          sheet::getMergedRegions);
    }
  }

  /** An XLSX workbook, each sheet streamed from its part. */
  static void parse(XlsxSheets workbook, ParseOptions options, Consumer<ParseEvent> emit,
                    ParseStatus.Builder status) throws IOException {
    workbook.defineNames(status);
    Conversion conversion = new Conversion(options, emit, status, UNBATCHED_SHEET_BYTES);
    int index = 0;
    for (XlsxSheets.Worksheet sheet : workbook.sheets()) {
      try (XlsxSheets.Rows rows = XlsxSheets.rows(sheet, status)) {
        conversion.sheet(index++, sheet.name(), sheet.hidden(), rows,
            () -> conversion.ranges(sheet.name(), rows.mergedReferences()));
        if (rows.skippedRows() > 0) {
          DocumentFaults.warn(status, "sheet '" + sheet.name() + "' has " + rows.skippedRows()
              + " rows numbered outside 1 to 1048576; they were skipped");
        }
      }
    }
  }

  /** Per-workbook conversion state: the formatter, the text budget, the batching choice. */
  private static final class Conversion {
    private final DataFormatter formatter = new DataFormatter();
    // A shared string costs its length every time a cell repeats it, so the
    // cells are charged for what they yield, not for what the file stores.
    private final TextBudget budget = new TextBudget("spreadsheet text");
    private final boolean batching;
    private final long unbatchedSheetBytes;
    private final Consumer<ParseEvent> emit;
    private final ParseStatus.Builder status;

    Conversion(ParseOptions options, Consumer<ParseEvent> emit, ParseStatus.Builder status,
               long unbatchedSheetBytes) {
      this.batching = options.sheetBatches();
      this.unbatchedSheetBytes = unbatchedSheetBytes;
      this.emit = emit;
      this.status = status;
    }

    /**
     * Converts one sheet. Merged regions are asked for only after the rows
     * run out, because a streamed sheet declares them after its rows.
     */
    void sheet(int index, String name, boolean hidden, Iterator<? extends Row> rows,
               Supplier<List<CellRangeAddress>> mergedRegions) {
      Batches batches = new Batches(index, name, hidden);
      while (rows.hasNext()) {
        Row row = rows.next();
        SheetRow.Builder convertedRow = SheetRow.newBuilder().setRowIndex(row.getRowNum());
        for (Cell cell : row) {
          SheetCell.Builder convertedCell;
          try {
            convertedCell = convert(cell, formatter);
          } catch (RuntimeException error) {
            // One unreadable cell (an unknown error code, a non-numeric
            // number) costs that cell, not the workbook.
            DocumentFaults.skip(status, "sheet '" + name + "' cell "
                + new CellAddress(cell).formatAsString(), error);
            continue;
          }
          if (convertedCell == null) continue;
          budget.spend(CELL_COST + convertedCell.getFormatted().length()
              + convertedCell.getFormula().length()
              + (convertedCell.hasText() ? convertedCell.getText().length() : 0));
          convertedRow.addCells(convertedCell);
        }
        if (convertedRow.getCellsCount() > 0) batches.add(convertedRow.build());
      }
      batches.finish(merged(name, mergedRegions.get()));
      status.setSheets(status.getSheets() + 1);
    }

    /** Parses streamed merged-cell references; an unreadable one is skipped. */
    List<CellRangeAddress> ranges(String sheetName, List<String> references) {
      List<CellRangeAddress> ranges = new ArrayList<>(references.size());
      for (String reference : references) {
        try {
          ranges.add(CellRangeAddress.valueOf(reference));
        } catch (RuntimeException error) {
          DocumentFaults.skip(status, "sheet '" + sheetName + "' merged region", error);
        }
      }
      return ranges;
    }

    private List<CellRange> merged(String sheetName, List<CellRangeAddress> regions) {
      List<CellRange> merged = new ArrayList<>();
      for (CellRangeAddress region : regions) {
        if (merged.size() == MAX_MERGED_REGIONS) {
          DocumentFaults.warn(status, "sheet '" + sheetName + "' has more than "
              + MAX_MERGED_REGIONS + " merged regions; the rest were skipped");
          break;
        }
        int firstRow = Math.min(region.getFirstRow(), region.getLastRow());
        int lastRow = Math.max(region.getFirstRow(), region.getLastRow());
        int firstColumn = Math.min(region.getFirstColumn(), region.getLastColumn());
        int lastColumn = Math.max(region.getFirstColumn(), region.getLastColumn());
        if (firstRow < 0 || firstColumn < 0) {
          // A whole-row or whole-column reference names no cells to anchor.
          DocumentFaults.warn(status, "sheet '" + sheetName + "' merged region "
              + region.formatAsString() + " is unbounded and was skipped");
          continue;
        }
        merged.add(CellRange.newBuilder()
            .setFirstRow(firstRow).setLastRow(lastRow)
            .setFirstColumn(firstColumn).setLastColumn(lastColumn)
            .build());
      }
      return merged;
    }

    /**
     * One sheet's rows as Sheet events. Without the client's opt-in a sheet
     * is a single event, which is what a consumer that predates batching
     * reads; one past the unbatched cap fails the parse instead of growing
     * without bound. With it, events of about
     * {@link #BATCH_BYTES}; a full batch is held until the next row arrives,
     * so the last event is never an empty one and is always the one with
     * more_rows unset.
     */
    private final class Batches {
      private final int index;
      private final String name;
      private final boolean hidden;
      private ai.pipestream.poi.v1.Sheet.Builder current;
      private long currentBytes;
      private ai.pipestream.poi.v1.Sheet.Builder full;

      Batches(int index, String name, boolean hidden) {
        this.index = index;
        this.name = name;
        this.hidden = hidden;
        this.current = open();
      }

      void add(SheetRow row) {
        if (full != null) {
          emit.accept(ParseEvent.newBuilder().setSheet(full.setMoreRows(true)).build());
          full = null;
        }
        currentBytes += CodedOutputStream.computeMessageSize(
            ai.pipestream.poi.v1.Sheet.ROWS_FIELD_NUMBER, row);
        if (!batching && currentBytes > unbatchedSheetBytes) {
          throw new DocumentTooLargeException("sheet '" + name + "' exceeds "
              + unbatchedSheetBytes + " bytes as one event; set sheet_batches to receive it"
              + " in batches");
        }
        current.addRows(row);
        if (batching && currentBytes >= BATCH_BYTES) {
          full = current;
          current = open();
          currentBytes = 0;
        }
      }

      void finish(List<CellRange> merged) {
        ai.pipestream.poi.v1.Sheet.Builder last = full != null ? full : current;
        emit.accept(ParseEvent.newBuilder()
            .setSheet(last.addAllMergedRegions(merged).setMoreRows(false))
            .build());
      }

      private ai.pipestream.poi.v1.Sheet.Builder open() {
        return ai.pipestream.poi.v1.Sheet.newBuilder()
            .setIndex(index).setName(name).setHidden(hidden);
      }
    }
  }

  private static SheetCell.Builder convert(Cell cell, DataFormatter formatter) {
    SheetCell.Builder converted = SheetCell.newBuilder().setColumnIndex(cell.getColumnIndex());
    CellType type = cell.getCellType();
    if (type == CellType.FORMULA) {
      converted.setFormula(cell.getCellFormula());
      type = cell.getCachedFormulaResultType();
      converted.setFormatted(cachedFormatted(cell, type));
    } else if (type != CellType.BLANK) {
      converted.setFormatted(formatter.formatCellValue(cell));
    }
    switch (type) {
      case STRING -> converted.setText(cell.getStringCellValue());
      case BOOLEAN -> converted.setBoolean(cell.getBooleanCellValue());
      case NUMERIC -> {
        if (DateUtil.isCellDateFormatted(cell)) {
          converted.setDate(ProtoTimestamps.fromDate(cell.getDateCellValue()));
        } else {
          converted.setNumber(cell.getNumericCellValue());
        }
      }
      case ERROR -> converted.setText(FormulaError.forInt(cell.getErrorCellValue()).getString());
      case BLANK -> {
        return null;
      }
      default -> {
        return null;
      }
    }
    return converted;
  }

  private static String cachedFormatted(Cell cell, CellType cachedType) {
    return switch (cachedType) {
      case STRING -> cell.getStringCellValue();
      case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
      case NUMERIC -> String.valueOf(cell.getNumericCellValue());
      case ERROR -> FormulaError.forInt(cell.getErrorCellValue()).getString();
      default -> "";
    };
  }
}
