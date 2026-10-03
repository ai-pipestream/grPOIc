package ai.pipestream.grpoic.parse;

import java.util.function.Consumer;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtRow;

/**
 * Children of a WordprocessingML container in document order, with content
 * controls (w:sdt) flattened away.
 *
 * <p><b>Why a walker of our own.</b> A content control may wrap a block
 * (paragraphs and tables), a table row, or a table cell, and XWPF only
 * models the first. Its row and cell lists skip wrapped rows and cells, so a
 * repeating-section control would drop rows and a wrapped cell would shift
 * every later column left. Walking the XML sees what Word renders.
 *
 * <p><b>Iterative on purpose.</b> Controls nest, and a crafted document can
 * nest them thousands deep; a cursor walk keeps its depth in a counter
 * instead of on the stack.
 */
final class ContentControls {

  private ContentControls() {}

  /**
   * Visits every child of {@code container} that is not itself a content
   * control or control body, descending into those instead. The control's
   * own properties (w:sdtPr, w:sdtEndPr) are visited like any other child;
   * callers ignore types they do not handle.
   */
  static void forEachChild(XmlObject container, Consumer<XmlObject> visit) {
    try (XmlCursor cursor = container.newCursor()) {
      if (!cursor.toFirstChild()) return;
      int depth = 0;
      while (true) {
        XmlObject child = cursor.getObject();
        boolean wrapper = isWrapper(child);
        if (wrapper && cursor.toFirstChild()) {
          depth++;
          continue;
        }
        if (!wrapper) visit.accept(child);
        while (!cursor.toNextSibling()) {
          if (depth == 0) return;
          cursor.toParent();
          depth--;
        }
      }
    }
  }

  private static boolean isWrapper(XmlObject element) {
    return element instanceof CTSdtBlock || element instanceof CTSdtContentBlock
        || element instanceof CTSdtRow || element instanceof CTSdtContentRow
        || element instanceof CTSdtCell || element instanceof CTSdtContentCell;
  }
}
