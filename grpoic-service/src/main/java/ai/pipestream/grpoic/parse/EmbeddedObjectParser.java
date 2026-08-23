package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.EmbeddedObject;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.List;
import java.util.function.Consumer;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.opc.PackagePart;

/**
 * Embedded parts an OOXML document carries, as descriptors only; bytes are
 * not streamed in v1. A listing failure degrades the parse to PARTIAL rather
 * than failing it.
 */
final class EmbeddedObjectParser {

  private EmbeddedObjectParser() {}

  static void parse(POIXMLDocument document, Consumer<ParseEvent> emit,
                    ParseStatus.Builder status) {
    List<PackagePart> parts;
    try {
      parts = document.getAllEmbeddedParts();
    } catch (Exception error) {
      status.setState(ParseStatus.State.STATE_PARTIAL);
      status.addWarnings("embedded part listing failed: " + error.getMessage());
      return;
    }
    for (PackagePart part : parts) {
      String name = part.getPartName().getName();
      emit.accept(
          ParseEvent.newBuilder()
              .setEmbeddedObject(
                  EmbeddedObject.newBuilder()
                      .setId(name)
                      .setFilename(name.substring(name.lastIndexOf('/') + 1))
                      .setContentType(part.getContentType())
                      .setSizeBytes(Math.max(part.getSize(), 0)))
              .build());
      status.setEmbeddedObjects(status.getEmbeddedObjects() + 1);
    }
  }
}
