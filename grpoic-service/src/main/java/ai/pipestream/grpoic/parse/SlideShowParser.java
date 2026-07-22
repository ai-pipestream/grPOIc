package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.function.Consumer;
import org.apache.poi.sl.usermodel.Notes;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.TextShape;

/**
 * Presentations through the common SL interface, so XSLF (.pptx) and HSLF
 * (.ppt) share one path.
 */
final class SlideShowParser {

  private SlideShowParser() {}

  static void parse(SlideShow<?, ?> show, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    int index = 0;
    for (Slide<?, ?> slide : show.getSlides()) {
      ai.pipestream.poi.v1.Slide.Builder converted =
          ai.pipestream.poi.v1.Slide.newBuilder().setIndex(index++);
      String title = slide.getTitle();
      if (title != null) converted.setTitle(title);
      for (Shape<?, ?> shape : slide.getShapes()) {
        if (!(shape instanceof TextShape<?, ?> textShape)) continue;
        if (isTitle(textShape)) continue;
        String text = textShape.getText();
        if (text != null && !text.isBlank()) converted.addTexts(text);
      }
      Notes<?, ?> notes = slide.getNotes();
      if (notes != null) {
        for (Shape<?, ?> shape : notes.getShapes()) {
          if (!(shape instanceof TextShape<?, ?> textShape)) continue;
          String text = textShape.getText();
          if (text != null && !text.isBlank()) converted.addNotes(text);
        }
      }
      emit.accept(ParseEvent.newBuilder().setSlide(converted).build());
      status.setSlides(status.getSlides() + 1);
    }
  }

  private static boolean isTitle(TextShape<?, ?> shape) {
    Placeholder placeholder = shape.getPlaceholder();
    return placeholder == Placeholder.TITLE || placeholder == Placeholder.CENTERED_TITLE;
  }
}
