package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.Paragraph;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.io.IOException;
import java.util.function.Consumer;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

/**
 * Legacy DOC body text. HWPF has no body-element iteration; WordExtractor's
 * paragraph pass keeps text order and strips field/cell control characters.
 * Tables arrive as text, which the status reports as a degradation.
 */
final class LegacyWordParser {

  private LegacyWordParser() {}

  static void parse(HWPFDocument document, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    try (WordExtractor extractor = new WordExtractor(document)) {
      for (String paragraph : extractor.getParagraphText()) {
        String text = WordExtractor.stripFields(paragraph).strip();
        if (text.isEmpty()) continue;
        emit.accept(
            ParseEvent.newBuilder().setParagraph(Paragraph.newBuilder().setText(text)).build());
        status.setParagraphs(status.getParagraphs() + 1);
      }
    } catch (IOException error) {
      throw new InvalidDocumentException("legacy DOC text extraction failed", error);
    }
    status.setState(ParseStatus.State.STATE_PARTIAL);
    status.addWarnings("legacy DOC tables and embedded objects are emitted as plain text");
  }
}
