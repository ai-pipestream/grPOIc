package ai.pipestream.grpoic.parse;

import java.io.IOException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFAbstractFootnoteEndnote;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFEndnote;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * An {@link XWPFDocument} that charges the text a footnote or endnote
 * reference pulls into a paragraph against the installed text limit.
 *
 * <p><b>Why.</b> XWPF resolves a note reference while it builds the
 * paragraph: it looks the note up on the document, renders the note's
 * paragraphs and appends that text to the paragraph's own, which keeps it.
 * A note may itself hold references, so a few bytes of XML describe a
 * note whose text is ten copies of the one below it, and a short chain
 * renders to gigabytes before the first event is emitted. The lookup is
 * the one seam the usermodel exposes, so the charge is taken there, before
 * the text is built.
 *
 * <p>The budget is created on first use: XWPFDocument's constructor reads
 * the document, and with it the first references, before this class's
 * fields exist.
 */
final class BudgetedWordDocument extends XWPFDocument {

  private TextBudget budget;

  BudgetedWordDocument(OPCPackage container) throws IOException {
    super(container);
  }

  @Override
  public XWPFFootnote getFootnoteByID(int id) {
    return charge(super.getFootnoteByID(id));
  }

  @Override
  public XWPFEndnote getEndnoteByID(int id) {
    return charge(super.getEndnoteByID(id));
  }

  /** Spends the note's text, which the caller is about to copy into a paragraph. */
  private <N extends XWPFAbstractFootnoteEndnote> N charge(N note) {
    if (note == null) return null;
    if (budget == null) budget = new TextBudget("note text");
    for (XWPFParagraph paragraph : note.getParagraphs()) {
      budget.spend(paragraph.getText().length());
    }
    return note;
  }
}
