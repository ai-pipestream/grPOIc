package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetCell;
import ai.pipestream.poi.v1.SheetRow;
import java.io.IOException;
import java.util.Iterator;
import java.util.function.Consumer;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

/**
 * Spreadsheets through the common SS interface, so XSSF (.xlsx) and HSSF
 * (.xls) share one cell path. XLSX sheets stream a row at a time
 * ({@link XlsxSheets}); XLS, whose format caps a sheet at 65,536 rows, goes
 * through HSSF's usermodel. Formula cells are never evaluated; they carry
 * the formula source and the workbook's cached result.
 */
final class SpreadsheetParser {

  private SpreadsheetParser() {}

  /** A workbook loaded whole by its usermodel (XLS). */
  static void parse(Workbook workbook, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    Conversion conversion = new Conversion(emit, status);
    for (int index = 0; index < workbook.getNumberOfSheets(); index++) {
      Sheet sheet = workbook.getSheetAt(index);
      conversion.sheet(index, sheet.getSheetName(), sheet.iterator());
    }
  }

  /** An XLSX workbook, each sheet streamed from its part. */
  static void parse(XlsxSheets workbook, Consumer<ParseEvent> emit, ParseStatus.Builder status)
      throws IOException {
    Conversion conversion = new Conversion(emit, status);
    int index = 0;
    for (XlsxSheets.Worksheet sheet : workbook.sheets()) {
      try (XlsxSheets.Rows rows = XlsxSheets.rows(sheet)) {
        conversion.sheet(index++, sheet.name(), rows);
      }
    }
  }

  /** Per-workbook conversion state: the formatter and the text budget. */
  private static final class Conversion {
    private final DataFormatter formatter = new DataFormatter();
    private final TextBudget budget = new TextBudget(ZipSecureFile.getMaxTextSize());
    private final Consumer<ParseEvent> emit;
    private final ParseStatus.Builder status;

    Conversion(Consumer<ParseEvent> emit, ParseStatus.Builder status) {
      this.emit = emit;
      this.status = status;
    }

    void sheet(int index, String name, Iterator<? extends Row> rows) {
      ai.pipestream.poi.v1.Sheet.Builder converted =
          ai.pipestream.poi.v1.Sheet.newBuilder().setIndex(index).setName(name);
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
          budget.spend(convertedCell);
          convertedRow.addCells(convertedCell);
        }
        if (convertedRow.getCellsCount() > 0) converted.addRows(convertedRow);
      }
      emit.accept(ParseEvent.newBuilder().setSheet(converted).build());
      status.setSheets(status.getSheets() + 1);
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

  /**
   * Characters of cell text the workbook may still yield. A shared string
   * costs its length every time a cell repeats it, so a small file cannot
   * expand into an unbounded stream.
   */
  private static final class TextBudget {
    private final long limit;
    private long spent;

    TextBudget(long limit) {
      this.limit = limit;
    }

    void spend(SheetCell.Builder cell) {
      spent += cell.getFormatted().length() + cell.getFormula().length()
          + (cell.hasText() ? cell.getText().length() : 0);
      if (spent > limit) {
        throw new DocumentTooLargeException(
            "spreadsheet text exceeds " + limit + " characters");
      }
    }
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
