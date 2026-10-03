package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/**
 * Without sheet batches a sheet is one event, so its rows pile up in one
 * builder. Past the unbatched cap the parse fails with a hint to batch,
 * and a batching client gets the same sheet in full.
 */
class UnbatchedSheetCapTest {

  private static final long CAP = 4096;

  /** A small first sheet, then one whose rows add up to well over {@link #CAP}. */
  private static XSSFWorkbook workbook() {
    XSSFWorkbook workbook = new XSSFWorkbook();
    workbook.createSheet("Small").createRow(0).createCell(0).setCellValue("fits");
    XSSFSheet wide = workbook.createSheet("Wide");
    for (int number = 0; number < 1000; number++) {
      Row row = wide.createRow(number);
      for (int column = 0; column < 4; column++) row.createCell(column).setCellValue(number);
    }
    return workbook;
  }

  @Test
  void unbatchedSheetPastTheCapFailsWithAHintToBatch() throws Exception {
    List<ParseEvent> events = new ArrayList<>();
    try (XSSFWorkbook workbook = workbook()) {
      assertThatThrownBy(() -> SpreadsheetParser.parse(workbook, ParseOptions.DEFAULTS,
          events::add, ParseStatus.newBuilder(), CAP))
          .isInstanceOf(DocumentTooLargeException.class)
          .hasMessageContaining("sheet 'Wide' exceeds " + CAP + " bytes")
          .hasMessageContaining("set sheet_batches");
    }
    assertThat(events).extracting(event -> event.getSheet().getName())
        .as("the sheet under the cap went out; the one over it did not")
        .containsExactly("Small");
  }

  @Test
  void batchedSheetIsNotSubjectToTheCap() throws Exception {
    List<ParseEvent> events = new ArrayList<>();
    ParseStatus.Builder status = ParseStatus.newBuilder();
    try (XSSFWorkbook workbook = workbook()) {
      SpreadsheetParser.parse(workbook, new ParseOptions(true), events::add, status, CAP);
    }
    assertThat(events.stream().filter(event -> event.getSheet().getName().equals("Wide"))
        .mapToInt(event -> event.getSheet().getRowsCount()).sum()).isEqualTo(1000);
    assertThat(status.getSheets()).isEqualTo(2);
  }
}
