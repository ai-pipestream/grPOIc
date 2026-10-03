package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Supplier;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

/**
 * DOCX body: paragraphs and tables interleaved in document order. A
 * block-level content control (w:sdt: a table of contents, a template's
 * fill-in region) is opened up, and its paragraphs and tables are emitted
 * in place as if they sat in the body.
 */
final class WordParser {

  private WordParser() {}

  static void parse(XWPFDocument document, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    // XWPF models a block content control as one opaque XWPFSDT; the body's
    // w:sdt elements, in the same order, give access to what they hold.
    List<CTSdtBlock> controls = document.getDocument().getBody().getSdtList();
    int nextControl = 0;
    for (IBodyElement element : document.getBodyElements()) {
      if (element instanceof XWPFParagraph paragraph) {
        paragraph(() -> paragraph, emit, status);
      } else if (element instanceof XWPFTable table) {
        table(() -> table, emit, status);
      } else if (element instanceof XWPFSDT && nextControl < controls.size()) {
        ContentControls.forEachChild(controls.get(nextControl++), child -> {
          if (child instanceof CTP paragraph) {
            paragraph(() -> new XWPFParagraph(paragraph, document), emit, status);
          } else if (child instanceof CTTbl table) {
            table(() -> new XWPFTable(table, document), emit, status);
          }
        });
      }
    }
  }

  // Conversion failures caused by the document skip that one block with a
  // warning; the emit stays outside the guard so a failing client is never
  // mistaken for a broken paragraph.
  private static void paragraph(Supplier<XWPFParagraph> source, Consumer<ParseEvent> emit,
                                ParseStatus.Builder status) {
    ai.pipestream.poi.v1.Paragraph converted;
    try {
      XWPFParagraph paragraph = source.get();
      String text = paragraph.getText();
      if (text == null || text.isBlank()) return;
      converted = ai.pipestream.poi.v1.Paragraph.newBuilder()
          .setText(text)
          .setStyle(paragraph.getStyle() == null ? "" : paragraph.getStyle())
          .build();
    } catch (RuntimeException error) {
      DocumentFaults.skip(status, "a paragraph", error);
      return;
    }
    emit.accept(ParseEvent.newBuilder().setParagraph(converted).build());
    status.setParagraphs(status.getParagraphs() + 1);
  }

  private static void table(Supplier<XWPFTable> source, Consumer<ParseEvent> emit,
                            ParseStatus.Builder status) {
    ai.pipestream.poi.v1.Table converted;
    try {
      converted = WordTables.convert(source.get());
    } catch (RuntimeException error) {
      DocumentFaults.skip(status, "a table", error);
      return;
    }
    emit.accept(ParseEvent.newBuilder().setTable(converted).build());
    status.setTables(status.getTables() + 1);
  }
}
