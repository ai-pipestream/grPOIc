package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.poi.v1.ParseStatus;
import org.apache.poi.util.RecordFormatException;
import org.junit.jupiter.api.Test;

/**
 * The document-or-server decision behind INVALID_ARGUMENT versus INTERNAL,
 * pinned on hand-built stack traces so it cannot drift silently.
 */
class DocumentFaultsTest {

  private static RuntimeException thrownFrom(RuntimeException error, String... classes) {
    StackTraceElement[] frames = new StackTraceElement[classes.length];
    for (int index = 0; index < classes.length; index++) {
      frames[index] = new StackTraceElement(classes[index], "read", null, 1);
    }
    error.setStackTrace(frames);
    return error;
  }

  @Test
  void runtimeErrorsThrownInsideParserLibrariesAreTheDocumentsFault() {
    assertThat(DocumentFaults.fromDocument(thrownFrom(new IllegalStateException(),
        "org.apache.poi.xssf.usermodel.XSSFCell", "ai.pipestream.grpoic.parse.SpreadsheetParser")))
        .isTrue();
    assertThat(DocumentFaults.fromDocument(thrownFrom(new ArrayIndexOutOfBoundsException(),
        "java.util.ArrayList", "jdk.internal.util.Preconditions",
        "org.apache.xmlbeans.impl.store.Cur", "ai.pipestream.grpoic.parse.WordParser")))
        .as("JDK frames are looked through to the library that called them")
        .isTrue();
  }

  @Test
  void runtimeErrorsThrownInGrpoicCodeStayServerFaults() {
    assertThat(DocumentFaults.fromDocument(thrownFrom(new NullPointerException(),
        "ai.pipestream.grpoic.parse.WordTables", "org.apache.poi.xwpf.usermodel.XWPFTable")))
        .isFalse();
    assertThat(DocumentFaults.fromDocument(thrownFrom(new IllegalStateException(),
        "java.util.Objects", "io.grpc.internal.ServerCallImpl")))
        .as("a transport failure is never the document's fault")
        .isFalse();
  }

  @Test
  void poiCorruptionTypesAreTheDocumentsFaultWhereverThrown() {
    assertThat(DocumentFaults.fromDocument(
        thrownFrom(new RecordFormatException("bad record"), "ai.pipestream.grpoic.Anything")))
        .isTrue();
  }

  @Test
  void skipWarnsForDocumentFaultsAndRethrowsServerFaults() {
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    DocumentFaults.skip(status, "cell A1",
        thrownFrom(new IllegalArgumentException("Unknown error code"), "org.apache.poi.ss.X"));
    assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(status.getWarningsList())
        .containsExactly("cell A1 skipped: IllegalArgumentException: Unknown error code");

    RuntimeException bug = thrownFrom(new NullPointerException(), "ai.pipestream.grpoic.X");
    assertThatThrownBy(() -> DocumentFaults.skip(status, "cell B1", bug)).isSameAs(bug);
  }

  @Test
  void warningsStayBounded() {
    ParseStatus.Builder status = ParseStatus.newBuilder();
    for (int index = 0; index < 500; index++) DocumentFaults.warn(status, "warning " + index);
    assertThat(status.getWarningsCount()).isEqualTo(DocumentFaults.MAX_WARNINGS + 1);
    assertThat(status.getWarnings(DocumentFaults.MAX_WARNINGS)).isEqualTo("further warnings omitted");
  }
}
