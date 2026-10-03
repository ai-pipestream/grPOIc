package ai.pipestream.grpoic;

import static ai.pipestream.grpoic.OoxmlFixtures.withPart;
import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.grpoic.ParseHarness.ParseResult;
import ai.pipestream.poi.v1.CellRange;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.Sheet;
import ai.pipestream.poi.v1.SheetRow;
import java.io.ByteArrayOutputStream;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Sheets on the wire: a large sheet arrives as ordered batches of bounded
 * size instead of one message that grows with the sheet, and merged cells
 * arrive as typed ranges on the sheet's last batch, for XLSX and XLS alike.
 */
class SheetStructureTest {

  private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  private static final int BATCH_BYTES = 1 << 20;

  private static ParseHarness harness;

  @BeforeAll
  static void start() throws Exception {
    harness = new ParseHarness(16L * 1024 * 1024, 4);
  }

  @AfterAll
  static void stop() throws Exception {
    harness.close();
  }

  private static byte[] xlsx(String... sheetNames) throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      for (String name : sheetNames) workbook.createSheet(name);
      workbook.write(out);
      return out.toByteArray();
    }
  }

  private static String sheetXml(String rows, String mergeCells) {
    return "<worksheet xmlns=\"" + NS + "\"><sheetData>" + rows + "</sheetData>" + mergeCells
        + "</worksheet>";
  }

  @Test
  void largeSheetArrivesInOrderedBoundedBatches() throws Exception {
    int rowCount = 30_000;
    StringBuilder rows = new StringBuilder();
    for (int row = 1; row <= rowCount; row++) {
      rows.append("<row r=\"").append(row).append("\">");
      for (char column = 'A'; column <= 'E'; column++) {
        rows.append("<c r=\"").append(column).append(row)
            .append("\" t=\"inlineStr\"><is><t>value ").append(row).append(column)
            .append("</t></is></c>");
      }
      rows.append("</row>");
    }
    byte[] bytes = withPart(xlsx("Big", "After"), "/xl/worksheets/sheet1.xml",
        sheetXml(rows.toString(), "<mergeCells count=\"1\"><mergeCell ref=\"A1:E1\"/></mergeCells>"));

    ParseResult result = harness.parseOk(bytes, "big-sheet");
    List<Sheet> batches = result.eventsOf(ParseEvent::hasSheet).stream()
        .map(ParseEvent::getSheet).filter(sheet -> sheet.getIndex() == 0).toList();
    assertThat(batches.size()).as("about 2.5 MB of rows cannot be one batch").isGreaterThan(1);
    for (int index = 0; index < batches.size(); index++) {
      Sheet batch = batches.get(index);
      boolean last = index == batches.size() - 1;
      assertThat(batch.getName()).isEqualTo("Big");
      assertThat(batch.getMoreRows()).as("every batch but the last says more follow")
          .isEqualTo(!last);
      assertThat(batch.getMergedRegionsCount()).as("merged regions ride the last batch")
          .isEqualTo(last ? 1 : 0);
      assertThat(batch.getSerializedSize()).isLessThan(BATCH_BYTES + 4096);
    }
    List<SheetRow> all = batches.stream().flatMap(batch -> batch.getRowsList().stream()).toList();
    assertThat(all).hasSize(rowCount);
    for (int row = 0; row < rowCount; row++) {
      assertThat(all.get(row).getRowIndex()).isEqualTo(row);
    }
    assertThat(all.get(rowCount - 1).getCells(4).getText()).isEqualTo("value " + rowCount + "E");

    List<Sheet> after = result.eventsOf(ParseEvent::hasSheet).stream()
        .map(ParseEvent::getSheet).filter(sheet -> sheet.getIndex() == 1).toList();
    assertThat(after).hasSize(1);
    assertThat(after.get(0).getMoreRows()).isFalse();
    assertThat(result.status().getSheets()).as("sheets are counted, not batches").isEqualTo(2);
  }

  @Test
  void xlsxMergedRegionsAreTypedRanges() throws Exception {
    byte[] bytes = withPart(xlsx("Report"), "/xl/worksheets/sheet1.xml", sheetXml(
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Quarterly report</t></is></c></row>"
            + "<row r=\"3\"><c r=\"B3\" t=\"inlineStr\"><is><t>Region</t></is></c></row>",
        "<mergeCells count=\"4\"><mergeCell ref=\"A1:F1\"/><mergeCell ref=\"B3:B5\"/>"
            + "<mergeCell ref=\"not a range\"/><mergeCell ref=\"C:D\"/></mergeCells>"));
    ParseResult result = harness.parseOk(bytes, "merged-xlsx");
    Sheet sheet = result.eventsOf(ParseEvent::hasSheet).get(0).getSheet();
    assertThat(sheet.getMergedRegionsList()).containsExactly(
        range(0, 0, 0, 5),
        range(2, 4, 1, 1));
    assertThat(result.status().getState())
        .as("unreadable and unbounded ranges are skipped with a warning")
        .isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(result.status().getWarningsList()).hasSize(2);
  }

  @Test
  void xlsMergedRegionsAreTypedRanges() throws Exception {
    byte[] bytes;
    try (HSSFWorkbook workbook = new HSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      var sheet = workbook.createSheet("Legacy");
      sheet.createRow(0).createCell(0).setCellValue("title");
      sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
      sheet.addMergedRegion(new CellRangeAddress(2, 6, 1, 2));
      workbook.write(out);
      bytes = out.toByteArray();
    }
    Sheet sheet = harness.parseOk(bytes, "merged-xls").eventsOf(ParseEvent::hasSheet).get(0)
        .getSheet();
    assertThat(sheet.getMergedRegionsList()).containsExactly(
        range(0, 0, 0, 3),
        range(2, 6, 1, 2));
  }

  private static CellRange range(int firstRow, int lastRow, int firstColumn, int lastColumn) {
    return CellRange.newBuilder().setFirstRow(firstRow).setLastRow(lastRow)
        .setFirstColumn(firstColumn).setLastColumn(lastColumn).build();
  }
}
