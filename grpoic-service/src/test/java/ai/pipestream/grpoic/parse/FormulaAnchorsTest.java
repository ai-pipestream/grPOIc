package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.SheetCell;
import com.google.protobuf.ByteString;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;

/**
 * Shared-formula masters and array-formula anchors are held only while a
 * later row can use them. Before, each one was copied into the sheet's
 * scratch rows and the sheet's private formula state, which grew with the
 * sheet and lived until the whole workbook was done.
 */
class FormulaAnchorsTest {

  private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  /** A one-sheet workbook whose sheet part is {@code sheetData}. */
  private static byte[] workbook(String sheetData) throws Exception {
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      workbook.createSheet("Data");
      workbook.write(out);
      base = out.toByteArray();
    }
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(base));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      PackagePart part =
          container.getPart(PackagingURIHelper.createPartName("/xl/worksheets/sheet1.xml"));
      try (OutputStream partOut = part.getOutputStream()) {
        partOut.write(("<worksheet xmlns=\"" + NS + "\"><sheetData>" + sheetData
            + "</sheetData></worksheet>").getBytes(StandardCharsets.UTF_8));
      }
      container.save(out);
      return out.toByteArray();
    }
  }

  /**
   * Rows 1..n: A holds the row number; B:C carry a two-row array formula
   * anchored on every odd row; D carries a three-row shared formula
   * mastered on every third row.
   */
  private static String formulaRows(int rows) {
    StringBuilder xml = new StringBuilder();
    for (int r = 1; r <= rows; r++) {
      xml.append("<row r=\"").append(r).append("\">");
      xml.append("<c r=\"A").append(r).append("\"><v>").append(r).append("</v></c>");
      if (r % 2 == 1) {
        xml.append("<c r=\"B").append(r).append("\"><f t=\"array\" ref=\"B").append(r)
            .append(":C").append(r + 1).append("\">A").append(r).append("*2</f><v>0</v></c>");
      } else {
        xml.append("<c r=\"B").append(r).append("\"><v>0</v></c>");
      }
      xml.append("<c r=\"C").append(r).append("\"><v>0</v></c>");
      int group = (r - 1) / 3;
      if ((r - 1) % 3 == 0) {
        xml.append("<c r=\"D").append(r).append("\"><f t=\"shared\" ref=\"D").append(r)
            .append(":D").append(r + 2).append("\" si=\"").append(group).append("\">A")
            .append(r).append("+1</f><v>0</v></c>");
      } else {
        xml.append("<c r=\"D").append(r).append("\"><f t=\"shared\" si=\"").append(group)
            .append("\"/><v>0</v></c>");
      }
      xml.append("</row>");
    }
    return xml.toString();
  }

  @Test
  void aLongSheetHoldsOnlyTheGroupsCrossingTheCurrentRow() throws Exception {
    int rows = 20_000;
    byte[] bytes = workbook(formulaRows(rows));
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    try (OPCPackage container = InMemoryPackages.open(ByteString.copyFrom(bytes))) {
      XlsxSheets workbook = XlsxSheets.open(container);
      XlsxSheets.Worksheet sheet = workbook.sheets().get(0);
      int seen = 0;
      int mostLive = 0;
      try (XlsxSheets.Rows stream = XlsxSheets.rows(sheet, status)) {
        while (stream.hasNext()) {
          Row row = stream.next();
          int r = row.getRowNum() + 1;
          mostLive = Math.max(mostLive, stream.liveFormulaGroups());
          int anchorRow = r % 2 == 1 ? r : r - 1;
          for (String column : List.of("B", "C")) {
            Cell covered = row.getCell(CellReference.convertColStringToIndex(column));
            assertThat(covered.getCellFormula()).as("%s%d", column, r)
                .isEqualTo("A" + anchorRow + "*2");
          }
          assertThat(row.getCell(3).getCellFormula()).as("D%d", r).isEqualTo("A" + r + "+1");
          seen++;
        }
        assertThat(sheet.scratch().getSharedFormula(0))
            .as("a passed group is dropped").isNull();
      }
      assertThat(seen).isEqualTo(rows);
      assertThat(mostLive).as("one array and one shared group cross any row").isLessThanOrEqualTo(2);
      assertThat(sheet.scratch().getPhysicalNumberOfRows())
          .as("the scratch sheet keeps no rows").isZero();
      assertThat(sheet.scratch().getSharedFormula((rows - 1) / 3))
          .as("the sheet's groups go when its rows close").isNull();
    }
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_OK);
  }

  @Test
  void sharedMastersPastTheColumnCountAreCappedWithOneWarning() {
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    FormulaAnchors anchors = new FormulaAnchors("Data", status);
    int si = 0;
    for (int rowIndex = 0; rowIndex < 2; rowIndex++) {
      CTRow row = CTRow.Factory.newInstance();
      row.setR(rowIndex + 1);
      for (int column = 0; column < FormulaAnchors.MAX_LIVE; column++) {
        CTCell cell = row.addNewC();
        cell.setR(new CellReference(rowIndex, column).formatAsString());
        CTCellFormula formula = cell.addNewF();
        formula.setT(STCellFormulaType.SHARED);
        // Every group claims the rest of the sheet, so none is ever passed.
        formula.setRef(new CellReference(rowIndex, column).formatAsString() + ":"
            + new CellReference(1_048_575, column).formatAsString());
        formula.setSi(si++);
        formula.setStringValue("1");
      }
      anchors.prepare(row, rowIndex);
    }
    assertThat(anchors.live()).isEqualTo(FormulaAnchors.MAX_LIVE);
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(status.getWarningsList()).containsExactly("sheet 'Data' has more than "
        + FormulaAnchors.MAX_LIVE + " shared formulas open at once; cells using the rest were"
        + " skipped");
    assertThat(anchors.shared(0)).isNotNull();
    assertThat(anchors.shared(FormulaAnchors.MAX_LIVE)).isNull();
  }

  @Test
  void unreadableGroupsCostTheirCellsNotTheWorkbook() throws Exception {
    // A shared master with a broken ref and one with an si POI cannot key
    // used to fail the whole workbook as it bound the row.
    byte[] bytes = workbook("<row r=\"1\"><c r=\"A1\"><v>7</v></c>"
        + "<c r=\"B1\"><f t=\"shared\" ref=\"B1:\" si=\"0\">A1+1</f><v>8</v></c>"
        + "<c r=\"C1\"><f t=\"shared\" ref=\"C1:C2\" si=\"4294967295\">A1+2</f><v>9</v></c>"
        + "<c r=\"D1\"><f t=\"array\" ref=\"D1:\">A1*3</f><v>21</v></c>"
        + "<c r=\"E1\"><f>A1*4</f><v>28</v></c></row>");
    List<ParseEvent> events = new ArrayList<>();
    DocumentParser.parse("broken-groups", ByteString.copyFrom(bytes), ParseOptions.DEFAULTS,
        events::add);
    ParseStatus status = events.get(events.size() - 1).getStatus();
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    List<SheetCell> cells = events.stream().filter(ParseEvent::hasSheet)
        .flatMap(event -> event.getSheet().getRowsList().stream())
        .flatMap(row -> row.getCellsList().stream()).toList();
    assertThat(cells).extracting(SheetCell::getColumnIndex).contains(0, 3, 4);
    assertThat(cells).filteredOn(cell -> cell.getColumnIndex() == 3).singleElement()
        .extracting(SheetCell::getFormula).isEqualTo("A1*3");
    assertThat(cells).filteredOn(cell -> cell.getColumnIndex() == 4).singleElement()
        .extracting(SheetCell::getFormula).isEqualTo("A1*4");
    assertThat(status.getWarningsList())
        .anyMatch(warning -> warning.startsWith("sheet 'Data' shared formula at B1 skipped"))
        .anyMatch(warning -> warning.startsWith("sheet 'Data' shared formula at C1 skipped"))
        .anyMatch(warning -> warning.startsWith("sheet 'Data' array formula at D1 skipped"));
  }
}
