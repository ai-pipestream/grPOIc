package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.TableCell;
import ai.pipestream.poi.v1.TableRow;
import java.util.function.Consumer;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/** DOCX body: paragraphs and tables interleaved in document order. */
final class WordParser {

  private WordParser() {}

  static void parse(XWPFDocument document, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    for (IBodyElement element : document.getBodyElements()) {
      if (element instanceof XWPFParagraph paragraph) {
        String text = paragraph.getText();
        if (text == null || text.isBlank()) continue;
        emit.accept(
            ParseEvent.newBuilder()
                .setParagraph(
                    ai.pipestream.poi.v1.Paragraph.newBuilder()
                        .setText(text)
                        .setStyle(paragraph.getStyle() == null ? "" : paragraph.getStyle()))
                .build());
        status.setParagraphs(status.getParagraphs() + 1);
      } else if (element instanceof XWPFTable table) {
        emit.accept(ParseEvent.newBuilder().setTable(convert(table)).build());
        status.setTables(status.getTables() + 1);
      }
    }
  }

  private static ai.pipestream.poi.v1.Table convert(XWPFTable table) {
    ai.pipestream.poi.v1.Table.Builder converted = ai.pipestream.poi.v1.Table.newBuilder();
    for (XWPFTableRow row : table.getRows()) {
      TableRow.Builder convertedRow = TableRow.newBuilder();
      for (XWPFTableCell cell : row.getTableCells()) {
        convertedRow.addCells(
            TableCell.newBuilder()
                .setText(cell.getText())
                .setRowSpan(1)
                .setColSpan(gridSpan(cell)));
      }
      converted.addRows(convertedRow);
    }
    return converted.build();
  }

  private static int gridSpan(XWPFTableCell cell) {
    var properties = cell.getCTTc().getTcPr();
    if (properties == null || !properties.isSetGridSpan()) return 1;
    return properties.getGridSpan().getVal().intValue();
  }
}
