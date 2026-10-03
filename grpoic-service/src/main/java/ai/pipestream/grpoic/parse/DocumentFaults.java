package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.ParseStatus;
import java.util.List;
import org.apache.poi.EmptyFileException;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.util.RecordFormatException;

/**
 * Tells a broken document from a broken server, and records content that
 * had to be skipped.
 *
 * <p><b>Why the stack decides.</b> POI reports most corruption as plain
 * runtime exceptions ({@code IllegalArgumentException},
 * {@code IndexOutOfBoundsException}, {@code ClassCastException}, ...) thrown
 * while it reads the bytes, and the same types thrown from grPOIc's own code
 * mean a grPOIc bug. So a runtime exception counts as the document's fault
 * when the first frame outside the JDK belongs to a parser library, or when
 * it is one of POI's dedicated corruption types. Everything else stays a
 * server fault, so a real bug still surfaces as INTERNAL and counts as
 * failed instead of hiding behind INVALID_ARGUMENT.
 */
final class DocumentFaults {

  /** Warnings past this many collapse into one closing note. */
  static final int MAX_WARNINGS = 20;

  private static final List<String> PARSER_PACKAGES = List.of(
      "org.apache.poi.", "org.apache.xmlbeans.", "org.openxmlformats.",
      "com.microsoft.schemas.", "org.etsi.", "org.w3.", "org.apache.commons.");

  private static final List<String> JDK_PACKAGES =
      List.of("java.", "javax.", "jdk.", "sun.", "com.sun.");

  private DocumentFaults() {}

  /** True when the failure says the bytes are bad rather than grPOIc. */
  static boolean fromDocument(Throwable error) {
    if (error instanceof RecordFormatException || error instanceof POIXMLException
        || error instanceof OpenXML4JRuntimeException || error instanceof EmptyFileException) {
      return true;
    }
    if (!(error instanceof RuntimeException)) return false;
    for (StackTraceElement frame : error.getStackTrace()) {
      String type = frame.getClassName();
      if (JDK_PACKAGES.stream().anyMatch(type::startsWith)) continue;
      return PARSER_PACKAGES.stream().anyMatch(type::startsWith);
    }
    return false;
  }

  /**
   * Records that {@code what} was skipped because of a document fault and
   * marks the parse partial. A server fault is rethrown instead: it must
   * not be laundered into a warning.
   */
  static void skip(ParseStatus.Builder status, String what, RuntimeException error) {
    if (!fromDocument(error)) throw error;
    warn(status, what + " skipped: " + describe(error));
  }

  /** Adds a warning and marks the parse partial; the list stays bounded. */
  static void warn(ParseStatus.Builder status, String warning) {
    status.setState(ParseStatus.State.STATE_PARTIAL);
    int count = status.getWarningsCount();
    if (count < MAX_WARNINGS) {
      status.addWarnings(warning);
    } else if (count == MAX_WARNINGS) {
      status.addWarnings("further warnings omitted");
    }
  }

  /** A short account of a failure: its type and the start of its message. */
  static String describe(Throwable error) {
    String message = error.getMessage();
    if (message == null || message.isBlank()) return error.getClass().getSimpleName();
    message = message.strip();
    if (message.length() > 160) message = message.substring(0, 160) + "...";
    return error.getClass().getSimpleName() + ": " + message;
  }
}
