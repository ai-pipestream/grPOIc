package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.EmbeddedObject;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.util.List;
import java.util.function.Consumer;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;

/**
 * Embedded parts an OOXML document carries, as descriptors only; bytes are
 * not streamed in v1.
 *
 * <p>A damaged package costs the listing or the one object, with a warning
 * and STATE_PARTIAL, never the document. Only document faults degrade this
 * way: POI's checked package exceptions, and runtime exceptions
 * {@link DocumentFaults} attributes to the document. A runtime exception
 * from grPOIc's own code and any {@code Error} propagate, so a bug still
 * fails the call as INTERNAL. Warnings name the exception type only: POI
 * messages can quote part names and content.
 */
final class EmbeddedObjectParser {

  private EmbeddedObjectParser() {}

  /** Lists a package's embedded parts; may fail on a damaged package. */
  @FunctionalInterface
  interface Listing {
    List<PackagePart> parts() throws OpenXML4JException;
  }

  static void parse(POIXMLDocument document, Consumer<ParseEvent> emit,
                    ParseStatus.Builder status) {
    parse(document::getAllEmbeddedParts, emit, status);
  }

  static void parse(Listing listing, Consumer<ParseEvent> emit, ParseStatus.Builder status) {
    List<PackagePart> parts;
    try {
      parts = listing.parts();
    } catch (OpenXML4JException unreadable) {
      DocumentFaults.warn(status,
          "embedded objects skipped: " + unreadable.getClass().getSimpleName());
      return;
    } catch (RuntimeException error) {
      DocumentFaults.skipByType(status, "embedded objects", error);
      return;
    }
    int ordinal = 0;
    for (PackagePart part : parts) {
      ordinal++;
      if (part == null) {
        DocumentFaults.warn(status, "an embedding that points at a missing part was skipped");
        continue;
      }
      EmbeddedObject.Builder descriptor;
      try {
        descriptor = describe(part);
      } catch (RuntimeException error) {
        DocumentFaults.skipByType(status, "embedded object " + ordinal, error);
        continue;
      }
      emit.accept(ParseEvent.newBuilder().setEmbeddedObject(descriptor).build());
      status.setEmbeddedObjects(status.getEmbeddedObjects() + 1);
    }
  }

  private static EmbeddedObject.Builder describe(PackagePart part) {
    String name = part.getPartName().getName();
    return EmbeddedObject.newBuilder()
        .setId(name)
        .setFilename(name.substring(name.lastIndexOf('/') + 1))
        .setContentType(part.getContentType())
        .setSizeBytes(Math.max(part.getSize(), 0));
  }
}
