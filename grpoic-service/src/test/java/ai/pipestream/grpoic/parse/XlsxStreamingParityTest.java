package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import com.google.protobuf.ByteString;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/**
 * The streaming XLSX path must put exactly what the usermodel path puts on
 * the wire. Each workbook here goes through both: DocumentParser, which
 * streams, and the shared converter over a fully loaded XSSFWorkbook, which
 * is what this service did before. Their Sheet events and status must be
 * identical, cell for cell.
 */
class XlsxStreamingParityTest {

  private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  private record Output(List<ParseEvent> sheets, ParseStatus status) {}

  private static Output streamed(byte[] bytes) {
    List<ParseEvent> events = new ArrayList<>();
    DocumentParser.parse("parity", ByteString.copyFrom(bytes), ParseOptions.DEFAULTS, events::add);
    return new Output(events.stream().filter(ParseEvent::hasSheet).toList(),
        events.get(events.size() - 1).getStatus());
  }

  private static Output loaded(byte[] bytes) throws Exception {
    List<ParseEvent> events = new ArrayList<>();
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
      SpreadsheetParser.parse(workbook, ParseOptions.DEFAULTS, events::add, status);
    }
    return new Output(events, status.build());
  }

  private static void assertParity(byte[] bytes) throws Exception {
    Output expected = loaded(bytes);
    Output actual = streamed(bytes);
    assertThat(actual.sheets()).isEqualTo(expected.sheets());
    assertThat(actual.status().getStateValue()).isEqualTo(expected.status().getStateValue());
    assertThat(actual.status().getWarningsList()).isEqualTo(expected.status().getWarningsList());
    assertThat(expected.sheets()).as("the fixture must exercise something").isNotEmpty();
  }

  private static byte[] write(XSSFWorkbook workbook) throws Exception {
    try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      workbook.write(out);
      return out.toByteArray();
    }
  }

  private static byte[] withParts(byte[] bytes, String... namesAndXml) throws Exception {
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(bytes));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      for (int index = 0; index < namesAndXml.length; index += 2) {
        PackagePart part =
            container.getPart(PackagingURIHelper.createPartName(namesAndXml[index]));
        try (OutputStream partOut = part.getOutputStream()) {
          partOut.write(namesAndXml[index + 1].getBytes(StandardCharsets.UTF_8));
        }
      }
      container.save(out);
      return out.toByteArray();
    }
  }

  @Test
  void typedCellsFormatsAndFormulasMatch() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet data = workbook.createSheet("Data");
      XSSFSheet other = workbook.createSheet("Other Sheet");
      workbook.createSheet("Empty");
      String[] formats = {"0.00", "0.00E00", "0.0%", "#,##0", "General", "yyyy-mm-dd hh:mm:ss",
          "m/d/yy", "h:mm AM/PM", "[$-409]mmmm d, yyyy", "0.000_);(0.000)", "@", "# ?/?"};
      double[] values = {2.675, 123456.789, 0.1 + 0.2, -1234.5, 1e-7, 45000.25, 0, 1e20};
      int rowIndex = 0;
      for (String format : formats) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        Row row = data.createRow(rowIndex++);
        for (int column = 0; column < values.length; column++) {
          row.createCell(column).setCellValue(values[column]);
          row.getCell(column).setCellStyle(style);
        }
      }
      Row strings = data.createRow(rowIndex++);
      strings.createCell(0).setCellValue("carriage _x000D_ return and _x005F_x0041_ escape");
      XSSFRichTextString rich = new XSSFRichTextString("bold then plain");
      XSSFFont bold = workbook.createFont();
      bold.setBold(true);
      rich.applyFont(0, 4, bold);
      strings.createCell(1).setCellValue(rich);
      strings.createCell(2).setCellValue(true);
      strings.createCell(3).setCellErrorValue(FormulaError.NA.getCode());
      strings.createCell(4)
          .setCellValue(new Date(Instant.parse("1999-12-31T23:59:59Z").toEpochMilli()));
      strings.createCell(5).setBlank();
      strings.getCell(5).setCellStyle(workbook.createCellStyle());
      var rate = workbook.createName();
      rate.setNameName("Rate");
      rate.setRefersToFormula("Data!$B$1");
      Row formulas = data.createRow(rowIndex + 5);
      formulas.createCell(0).setCellFormula("A1*2");
      formulas.createCell(1).setCellFormula("A" + (rowIndex) + "&\"x\"");
      formulas.createCell(2).setCellFormula("1/0");
      formulas.createCell(3).setCellFormula("'Other Sheet'!A1+1");
      formulas.createCell(4).setCellFormula("TRUE");
      formulas.createCell(5).setCellFormula("Rate*3");
      formulas.createCell(6).setCellFormula("SUM(A1:H3)");
      other.createRow(0).createCell(0).setCellValue(5);
      other.createRow(1048575).createCell(16383).setCellValue("far corner");
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      assertParity(write(workbook));
    }
  }

  @Test
  void sharedAndArrayFormulasInlineStringsAndMissingReferencesMatch() throws Exception {
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Main");
      workbook.createSheet("Lookup Table").createRow(0).createCell(0).setCellValue(3);
      workbook.getSheet("Main").createRow(0).createCell(0).setCellValue("shared string");
      var local = workbook.createName();
      local.setSheetIndex(0);
      local.setNameName("Factor");
      local.setRefersToFormula("Main!$A$2");
      var global = workbook.createName();
      global.setNameName("Offset");
      global.setRefersToFormula("10");
      base = write(workbook);
    }
    String sheet = "<worksheet xmlns=\"" + NS + "\"><sheetData>"
        + "<row r=\"2\"><c r=\"A2\"><v>1</v></c>"
        + "<c r=\"B2\"><f t=\"shared\" ref=\"B2:C4\" si=\"0\">"
        + "A2*Factor+$A$2+'Lookup Table'!$A$1</f>"
        + "<v>5</v></c><c r=\"C2\"><f t=\"shared\" si=\"0\"/><v>6</v></c>"
        + "<c r=\"D2\"><f t=\"array\" ref=\"D2:D3\">SUM(A2:A4*2)+Offset</f><v>22</v></c>"
        + "<c r=\"E2\" t=\"inlineStr\"><is><r><t>inline </t></r><r><t>rich</t></r></is></c></row>"
        + "<row r=\"3\"><c r=\"A3\"><v>2</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>9</v></c>"
        + "<c r=\"C3\"><f t=\"shared\" si=\"0\"/><v>10</v></c><c r=\"D3\"><v>22</v></c>"
        + "<c r=\"E3\" t=\"str\"><f>\"s\"&amp;A3</f><v>s2</v></c></row>"
        + "<row><c r=\"A4\"><v>3</v></c><c><f t=\"shared\" si=\"0\"/><v>13</v></c>"
        + "<c t=\"e\"><v>#REF!</v></c><c t=\"b\"><v>0</v></c><c t=\"s\"><v>0</v></c>"
        + "<c t=\"s\"><v>99</v></c></row>"
        + "<row r=\"6\"><c r=\"A6\" t=\"n\"/><c r=\"B6\"><f>A4+1</f></c><c r=\"C6\"><v></v></c>"
        + "<c r=\"D6\" t=\"e\"><f>NA()</f><v>#N/A</v></c><c r=\"E6\"><v>  42  </v></c></row>"
        + "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A2:A3\"/></mergeCells></worksheet>";
    assertParity(withParts(base, "/xl/worksheets/sheet1.xml", sheet));
  }

  @Test
  void formulaGroupsThatCloseAndReopenDownTheSheetMatch() throws Exception {
    // Groups end and new ones start as the rows advance, so the streaming
    // path drops passed groups; a master that is not its range's top left
    // and a covered cell with an empty formula of its own are included.
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Groups");
      workbook.createSheet("After").createRow(0).createCell(0).setCellValue(1);
      base = write(workbook);
    }
    StringBuilder rows = new StringBuilder();
    for (int r = 1; r <= 30; r++) {
      rows.append("<row r=\"").append(r).append("\"><c r=\"A").append(r).append("\"><v>")
          .append(r).append("</v></c>");
      if (r % 3 == 1) {
        rows.append("<c r=\"B").append(r).append("\"><f t=\"array\" ref=\"B").append(r)
            .append(":C").append(r + 2).append("\">A").append(r).append("+A").append(r + 1)
            .append("</f><v>1</v></c><c r=\"C").append(r).append("\"><f></f><v>1</v></c>");
      } else {
        rows.append("<c r=\"B").append(r).append("\"><v>1</v></c><c r=\"C").append(r)
            .append("\"><v>1</v></c>");
      }
      int group = (r - 1) / 4;
      if ((r - 1) % 4 == 1) {
        rows.append("<c r=\"D").append(r).append("\"><f t=\"shared\" ref=\"D").append(r - 1)
            .append(":E").append(r + 2).append("\" si=\"").append(group).append("\">$A")
            .append(r).append("*2</f><v>2</v></c>");
      } else if ((r - 1) % 4 != 0) {
        rows.append("<c r=\"D").append(r).append("\"><f t=\"shared\" si=\"").append(group)
            .append("\"/><v>2</v></c>");
      }
      rows.append("<c r=\"E").append(r).append("\"><v>3</v></c></row>");
    }
    String sheet = "<worksheet xmlns=\"" + NS + "\"><sheetData>" + rows
        + "</sheetData></worksheet>";
    assertParity(withParts(base, "/xl/worksheets/sheet1.xml", sheet));
  }

  @Test
  void phoneticRunsAndTheNineteenFourDateSystemMatch() throws Exception {
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.getCTWorkbook().getWorkbookPr().setDate1904(true);
      CellStyle date = workbook.createCellStyle();
      date.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
      var row = workbook.createSheet("Dates").createRow(0);
      row.createCell(0).setCellValue("placeholder");
      row.createCell(1).setCellValue(40000);
      row.getCell(1).setCellStyle(date);
      base = write(workbook);
    }
    String strings = "<sst xmlns=\"" + NS + "\" count=\"1\" uniqueCount=\"1\">"
        + "<si><r><t>漢字</t></r>"
        + "<rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh><phoneticPr fontId=\"1\"/></si></sst>";
    assertParity(withParts(base, "/xl/sharedStrings.xml", strings));
  }

  @Test
  void unreadableCellsDegradeTheSameWay() throws Exception {
    byte[] base;
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Broken");
      base = write(workbook);
    }
    String sheet = "<worksheet xmlns=\"" + NS + "\"><sheetData>"
        + "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#BOGUS!</v></c><c r=\"B1\"><v>abc</v></c>"
        + "<c r=\"C1\"><v>4</v></c></row></sheetData></worksheet>";
    assertParity(withParts(base, "/xl/worksheets/sheet1.xml", sheet));
  }
}
