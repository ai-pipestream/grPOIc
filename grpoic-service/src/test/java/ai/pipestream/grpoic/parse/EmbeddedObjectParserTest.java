package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.util.RecordFormatException;
import org.junit.jupiter.api.Test;

/**
 * The embedded-object listing degrades only on document faults. Before, it
 * caught every Exception: a grPOIc bug became a warning, and the warning
 * quoted the exception's message, which can carry part names and content.
 */
class EmbeddedObjectParserTest {

  private static final String SECRET = "payroll-2026-salaries";

  private record Result(List<ParseEvent> events, ParseStatus status) {}

  private static Result run(EmbeddedObjectParser.Listing listing) {
    List<ParseEvent> events = new ArrayList<>();
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    EmbeddedObjectParser.parse(listing, events::add, status);
    return new Result(events, status.build());
  }

  private static <T extends Throwable> T thrownFrom(T error, String type) {
    error.setStackTrace(new StackTraceElement[] {new StackTraceElement(type, "read", null, 1)});
    return error;
  }

  private static PackagePart part(OPCPackage container, String name, boolean broken)
      throws InvalidFormatException {
    return new PackagePart(container, PackagingURIHelper.createPartName(name),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      @Override
      public long getSize() {
        if (broken) throw new RecordFormatException("size of " + SECRET + " unreadable");
        return 42;
      }

      @Override
      protected InputStream getInputStreamImpl() {
        return InputStream.nullInputStream();
      }

      @Override
      protected OutputStream getOutputStreamImpl() {
        return OutputStream.nullOutputStream();
      }

      @Override
      public boolean save(OutputStream out) {
        return true;
      }

      @Override
      public boolean load(InputStream in) {
        return true;
      }

      @Override
      public void close() {}

      @Override
      public void flush() {}
    };
  }

  @Test
  void aDamagedPackageListingIsPartialAndNamesOnlyTheExceptionType() {
    Result result = run(() -> {
      throw new InvalidFormatException("bad relationship to /word/embeddings/" + SECRET);
    });
    assertThat(result.status().getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(result.status().getWarningsList())
        .containsExactly("embedded objects skipped: InvalidFormatException");
    assertThat(result.events()).isEmpty();
  }

  @Test
  void aPoiRuntimeFaultInTheListingIsPartial() {
    Result result = run(() -> {
      throw thrownFrom(new IllegalArgumentException("part name " + SECRET),
          "org.apache.poi.openxml4j.opc.PackagePartName");
    });
    assertThat(result.status().getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
    assertThat(result.status().getWarningsList())
        .containsExactly("embedded objects skipped: IllegalArgumentException");
  }

  @Test
  void aGrpoicBugInTheListingIsNotSwallowed() {
    assertThatThrownBy(() -> run(() -> {
      throw thrownFrom(new NullPointerException(), "ai.pipestream.grpoic.parse.XlsxSheets");
    })).isInstanceOf(NullPointerException.class);
  }

  @Test
  void errorsPropagate() {
    assertThatThrownBy(() -> run(() -> {
      throw new OutOfMemoryError("Java heap space");
    })).isInstanceOf(OutOfMemoryError.class);
    assertThatThrownBy(() -> run(() -> {
      throw new StackOverflowError();
    })).isInstanceOf(StackOverflowError.class);
  }

  @Test
  void oneUnreadableObjectCostsOnlyItself() throws Exception {
    OPCPackage container = OPCPackage.create(new ByteArrayOutputStream());
    try {
      PackagePart broken = part(container, "/xl/embeddings/" + SECRET + ".xlsx", true);
      PackagePart fine = part(container, "/xl/embeddings/chart.xlsx", false);
      Result result = run(() -> List.of(broken, fine));
      assertThat(result.status().getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
      assertThat(result.status().getWarningsList())
          .containsExactly("embedded object 1 skipped: RecordFormatException");
      assertThat(result.status().getEmbeddedObjects()).isEqualTo(1);
      assertThat(result.events()).singleElement()
          .satisfies(event -> {
            assertThat(event.getEmbeddedObject().getFilename()).isEqualTo("chart.xlsx");
            assertThat(event.getEmbeddedObject().getSizeBytes()).isEqualTo(42);
          });
    } finally {
      container.revert();
    }
  }
}
