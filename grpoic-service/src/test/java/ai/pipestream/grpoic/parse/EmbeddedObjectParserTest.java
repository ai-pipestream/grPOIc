package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import com.google.protobuf.ByteString;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URI;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.util.RecordFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

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

  /**
   * A blank DOCX or XLSX whose host part (the document or the one sheet)
   * embeds one real OLE part and links one object outside the package. The
   * XLSX sheet also points one embedding at a folder, which is not a part
   * name; XWPFDocument already rejects that as it loads, so only the
   * streamed XLSX listing can meet it.
   */
  private static byte[] withLinkedObject(boolean word) throws Exception {
    byte[] base;
    try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      if (word) {
        try (XWPFDocument document = new XWPFDocument()) {
          document.createParagraph().createRun().setText("body");
          document.write(out);
        }
      } else {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
          workbook.createSheet("Data").createRow(0).createCell(0).setCellValue(1);
          workbook.write(out);
        }
      }
      base = out.toByteArray();
    }
    String root = word ? "/word" : "/xl";
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(base));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      PackagePart host = container.getPart(PackagingURIHelper.createPartName(
          word ? "/word/document.xml" : "/xl/worksheets/sheet1.xml"));
      PackagePart embedded = container.createPart(
          PackagingURIHelper.createPartName(root + "/embeddings/oleObject1.bin"),
          "application/vnd.openxmlformats-officedocument.oleObject");
      try (OutputStream partOut = embedded.getOutputStream()) {
        partOut.write(new byte[] {1, 2, 3});
      }
      host.addRelationship(embedded.getPartName(), TargetMode.INTERNAL,
          POIXMLDocument.OLE_OBJECT_REL_TYPE);
      host.addExternalRelationship("file:///C:/reports/" + SECRET + ".xlsx",
          POIXMLDocument.OLE_OBJECT_REL_TYPE);
      if (!word) {
        host.addRelationship(new URI(root + "/embeddings/"), TargetMode.INTERNAL,
            POIXMLDocument.PACK_OBJECT_REL_TYPE, "rIdFolder");
      }
      container.save(out);
      return out.toByteArray();
    }
  }

  @ParameterizedTest
  @ValueSource(booleans = {true, false})
  void aLinkedObjectIsLeftOutAndABadTargetCostsOnlyItself(boolean word) throws Exception {
    // Before, POI resolved every relationship in one call, so the linked
    // object's absolute URI failed the listing and every embedding was lost.
    List<ParseEvent> events = new ArrayList<>();
    DocumentParser.parse("linked", ByteString.copyFrom(withLinkedObject(word)),
        ParseOptions.DEFAULTS, events::add);
    ParseStatus status = events.get(events.size() - 1).getStatus();
    assertThat(events).filteredOn(ParseEvent::hasEmbeddedObject).singleElement()
        .satisfies(event -> {
          assertThat(event.getEmbeddedObject().getFilename()).isEqualTo("oleObject1.bin");
          assertThat(event.getEmbeddedObject().getSizeBytes()).isEqualTo(3);
        });
    assertThat(status.getEmbeddedObjects()).isEqualTo(1);
    if (word) {
      assertThat(status.getState()).as("a link is not damage")
          .isEqualTo(ParseStatus.State.STATE_OK);
      assertThat(status.getWarningsList()).isEmpty();
    } else {
      assertThat(status.getState()).isEqualTo(ParseStatus.State.STATE_PARTIAL);
      assertThat(status.getWarningsList()).containsExactly(
          "an embedding that points at a missing or unreadable part was skipped");
    }
  }
}
