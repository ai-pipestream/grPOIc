package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetCell;
import ai.pipestream.poi.v1.SheetRow;
import java.util.function.Consumer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

/**
 * Spreadsheets through the common SS interface, so XSSF (.xlsx) and HSSF
 * (.xls) share one path. Formula cells are never evaluated; they carry the
 * formula source and the workbook's cached result.
 */
final class SpreadsheetParser {

  private SpreadsheetParser() {}

  static void parse(Workbook workbook, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    DataFormatter formatter = new DataFormatter();
    TextBudget budget = new TextBudget(ZipSecureFile.getMaxTextSize());
    for (int index = 0; index < workbook.getNumberOfSheets(); index++) {
      Sheet sheet = workbook.getSheetAt(index);
      ai.pipestream.poi.v1.Sheet.Builder converted =
          ai.pipestream.poi.v1.Sheet.newBuilder().setIndex(index).setName(sheet.getSheetName());
      for (Row row : sheet) {
        SheetRow.Builder convertedRow = SheetRow.newBuilder().setRowIndex(row.getRowNum());
        for (Cell cell : row) {
          SheetCell.Builder convertedCell;
          try {
            convertedCell = convert(cell, formatter);
          } catch (RuntimeException error) {
            // One unreadable cell (an unknown error code, a non-numeric
            // number) costs that cell, not the workbook.
            DocumentFaults.skip(status, "sheet '" + sheet.getSheetName() + "' cell "
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
