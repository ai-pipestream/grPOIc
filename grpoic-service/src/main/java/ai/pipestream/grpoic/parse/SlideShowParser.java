package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import ai.pipestream.poi.v1.Table;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Iterator;
import java.util.List;
import java.util.function.Consumer;
import org.apache.poi.sl.usermodel.GroupShape;
import org.apache.poi.sl.usermodel.Notes;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.sl.usermodel.ShapeContainer;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.sl.usermodel.TextShape;

/**
 * Presentations through the common SL interface, so XSLF (.pptx) and HSLF
 * (.ppt) share one path. Shapes inside groups count as the slide's own, in
 * shape order; native tables follow their slide as Table events.
 */
final class SlideShowParser {

  private SlideShowParser() {}

  static void parse(SlideShow<?, ?> show, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    int index = 0;
    for (Slide<?, ?> slide : show.getSlides()) {
      int slideIndex = index++;
      String where = "slide " + (slideIndex + 1);
      // A shape or table the document breaks costs that shape with a
      // warning, a slide that cannot be walked at all costs the slide; the
      // emits stay outside the guards.
      ai.pipestream.poi.v1.Slide.Builder converted =
          ai.pipestream.poi.v1.Slide.newBuilder().setIndex(slideIndex);
      List<TableShape<?, ?>> tables = new ArrayList<>();
      try {
        String title = slide.getTitle();
        if (title != null) converted.setTitle(title);
        for (Shape<?, ?> shape : shapes(slide)) {
          if (shape instanceof TableShape<?, ?> table) {
            tables.add(table);
          } else if (shape instanceof TextShape<?, ?> textShape) {
            addText(textShape, true, converted::addTexts, status, where);
          }
        }
        Notes<?, ?> notes = slide.getNotes();
        if (notes != null) {
          for (Shape<?, ?> shape : shapes(notes)) {
            if (shape instanceof TextShape<?, ?> textShape) {
              addText(textShape, false, converted::addNotes, status, where + " notes");
            }
          }
        }
      } catch (RuntimeException error) {
        DocumentFaults.skip(status, where, error);
        continue;
      }
      emit.accept(ParseEvent.newBuilder().setSlide(converted).build());
      status.setSlides(status.getSlides() + 1);
      for (TableShape<?, ?> table : tables) {
        Table convertedTable;
        try {
          convertedTable = SlideTables.convert(table).toBuilder().setSlideIndex(slideIndex).build();
        } catch (RuntimeException error) {
          DocumentFaults.skip(status, "a table on " + where, error);
          continue;
        }
        emit.accept(ParseEvent.newBuilder().setTable(convertedTable).build());
        status.setTables(status.getTables() + 1);
      }
    }
  }

  private static void addText(TextShape<?, ?> shape, boolean skipTitle, Consumer<String> add,
                              ParseStatus.Builder status, String where) {
    try {
      if (skipTitle && isTitle(shape)) return;
      String text = shape.getText();
      if (text != null && !text.isBlank()) add.accept(text);
    } catch (RuntimeException error) {
      DocumentFaults.skip(status, "a text shape on " + where, error);
    }
  }

  /**
   * Every shape of the container in shape order, with group shapes replaced
   * by their members. Tables are returned whole, never opened as groups
   * (HSLF models a table as a group of text boxes). Iterative, so nesting
   * depth costs no stack.
   */
  private static List<Shape<?, ?>> shapes(ShapeContainer<?, ?> container) {
    List<Shape<?, ?>> shapes = new ArrayList<>();
    Deque<Iterator<? extends Shape<?, ?>>> pending = new ArrayDeque<>();
    pending.push(container.getShapes().iterator());
    while (!pending.isEmpty()) {
      Iterator<? extends Shape<?, ?>> siblings = pending.peek();
      if (!siblings.hasNext()) {
        pending.pop();
        continue;
      }
      Shape<?, ?> shape = siblings.next();
      if (!(shape instanceof TableShape) && shape instanceof GroupShape<?, ?> group) {
        pending.push(group.getShapes().iterator());
      } else {
        shapes.add(shape);
      }
    }
    return shapes;
  }

  private static boolean isTitle(TextShape<?, ?> shape) {
    Placeholder placeholder = shape.getPlaceholder();
    return placeholder == Placeholder.TITLE || placeholder == Placeholder.CENTERED_TITLE;
  }
}
