package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.function.Consumer;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

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
        emit.accept(ParseEvent.newBuilder().setTable(WordTables.convert(table)).build());
        status.setTables(status.getTables() + 1);
      }
    }
  }
}
