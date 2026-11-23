package ai.pipestream.grpoic;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.SheetCell;
import ai.pipestream.poi.v1.SheetRow;
import java.io.ByteArrayOutputStream;
import java.time.Instant;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * A formula cell's {@code formatted} is what the workbook displays, the same
 * as for a cell that holds a plain value: the cached result under the
 * cell's own number format. Totals, ratios and dates are usually formulas,
 * so this is the display text a consumer most often needs. The typed value
 * keeps the raw cached result either way.
 */
class FormulaFormattingTest {

  private static final Date WHEN = Date.from(Instant.parse("2024-03-15T00:00:00Z"));

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(4L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  /**
   * Row 0 holds the inputs, row 1 one formula per column: a ratio shown as
   * a percentage, a total shown as currency, a date, a comparison, a text
   * result, an error, and a number under the General format.
   */
  private static byte[] workbook(Workbook workbook) throws Exception {
    try (workbook; ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("Formulas");
      Row inputs = sheet.createRow(0);
      inputs.createCell(0).setCellValue(4567);
      inputs.createCell(1).setCellValue(10000);
      inputs.createCell(2).setCellValue(1234.5);
      Cell date = inputs.createCell(3);
      date.setCellValue(WHEN);
      date.setCellStyle(style(workbook, "yyyy-mm-dd"));
      inputs.createCell(4).setCellValue("north");

      Row formulas = sheet.createRow(1);
      formula(formulas, 0, "A1/B1", style(workbook, "0.00%"));
      formula(formulas, 1, "C1*1", style(workbook, "\"$\"#,##0.00"));
      formula(formulas, 2, "D1+0", style(workbook, "yyyy-mm-dd"));
      formula(formulas, 3, "A1<B1", null);
      formula(formulas, 4, "E1&\" region\"", null);
      formula(formulas, 5, "1/0", null);
      formula(formulas, 6, "A1/4", null);
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.write(out);
      return out.toByteArray();
    }
  }

  private static CellStyle style(Workbook workbook, String format) {
    CellStyle style = workbook.createCellStyle();
    style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(format));
    return style;
  }

  private static void formula(Row row, int column, String formula, CellStyle style) {
    Cell cell = row.createCell(column);
    cell.setCellFormula(formula);
    if (style != null) cell.setCellStyle(style);
  }

  private static void assertFormulaCellsShowWhatTheWorkbookDisplays(byte[] bytes, String id)
      throws Exception {
    SheetRow row = harness.parseOk(bytes, id).eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet().getRows(1);

    SheetCell ratio = row.getCells(0);
    assertThat(ratio.getFormula()).isEqualTo("A1/B1");
    assertThat(ratio.getNumber()).as("the typed value keeps the raw cached result")
        .isEqualTo(0.4567);
    assertThat(ratio.getFormatted()).as("a ratio formula shows as the percentage it is styled as")
        .isEqualTo("45.67%");

    SheetCell total = row.getCells(1);
    assertThat(total.getNumber()).isEqualTo(1234.5);
    assertThat(total.getFormatted()).as("a total formula shows under its currency format")
        .isEqualTo("$1,234.50");

    SheetCell date = row.getCells(2);
    assertThat(date.hasDate()).as("a date formula is typed as a date").isTrue();
    assertThat(date.getDate().getSeconds()).isEqualTo(WHEN.getTime() / 1000);
    assertThat(date.getFormatted()).as("a date formula shows the date, not its serial number")
        .isEqualTo("2024-03-15");

    SheetCell comparison = row.getCells(3);
    assertThat(comparison.getBoolean()).isTrue();
    assertThat(comparison.getFormatted())
        .as("a boolean result shows as the workbook shows it, as for a plain boolean cell")
        .isEqualTo("TRUE");

    SheetCell text = row.getCells(4);
    assertThat(text.getText()).isEqualTo("north region");
    assertThat(text.getFormatted()).isEqualTo("north region");

    SheetCell error = row.getCells(5);
    assertThat(error.getText()).isEqualTo("#DIV/0!");
    assertThat(error.getFormatted()).isEqualTo("#DIV/0!");

    SheetCell general = row.getCells(6);
    assertThat(general.getNumber()).isEqualTo(1141.75);
    assertThat(general.getFormatted())
        .as("a General number shows as the same cell would without a formula")
        .isEqualTo("1141.75");
  }

  @Test
  void xlsxFormulaCellsShowTheirCachedResultUnderTheCellsNumberFormat() throws Exception {
    assertFormulaCellsShowWhatTheWorkbookDisplays(workbook(new XSSFWorkbook()), "formulas-xlsx");
  }

  @Test
  void xlsFormulaCellsShowTheirCachedResultUnderTheCellsNumberFormat() throws Exception {
    assertFormulaCellsShowWhatTheWorkbookDisplays(workbook(new HSSFWorkbook()), "formulas-xls");
  }

  @Test
  void aWholeNumberFormulaResultShowsWithoutATrailingDecimal() throws Exception {
    byte[] bytes;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      Row row = workbook.createSheet("Sum").createRow(0);
      row.createCell(0).setCellValue(2);
      row.createCell(1).setCellValue(3);
      row.createCell(2).setCellFormula("SUM(A1:B1)");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      workbook.write(out);
      bytes = out.toByteArray();
    }
    SheetCell sum = harness.parseOk(bytes, "sum").eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet().getRows(0).getCells(2);
    assertThat(sum.getNumber()).isEqualTo(5.0);
    assertThat(sum.getFormatted()).as("5, as the workbook shows it, not 5.0").isEqualTo("5");
  }
}
