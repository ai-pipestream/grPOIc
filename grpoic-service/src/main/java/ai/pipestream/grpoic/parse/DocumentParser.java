package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.DocumentInfo;
import ai.pipestream.poi.v1.DocumentMetadata;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.function.Consumer;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Detects the format from the bytes (the advisory content type is never
 * trusted) and streams typed events: DocumentInfo first, content blocks in
 * document order, ParseStatus last. Entirely in memory; nothing is written to
 * disk and no process is executed. Per-format extraction lives in the sibling
 * parser classes; this class only detects, dispatches, and frames the stream.
 */
public final class DocumentParser {

  private DocumentParser() {}

  public static void parse(String documentId, byte[] bytes, Consumer<ParseEvent> emit) {
    FileMagic magic = FileMagic.valueOf(bytes);
    try {
      switch (magic) {
        case OOXML -> parseOoxml(documentId, bytes, emit);
        case OLE2 -> parseOle2(documentId, bytes, emit);
        default -> throw new UnsupportedFormatException(
            "not an office document (magic: " + magic + ")");
      }
    } catch (IOException error) {
      throw new InvalidDocumentException("unreadable document: " + error.getMessage(), error);
    }
  }

  private static void parseOoxml(String documentId, byte[] bytes, Consumer<ParseEvent> emit)
      throws IOException {
    try (OPCPackage container = openPackage(bytes)) {
      String coreType = coreContentType(container);
      if (coreType.contains("wordprocessingml")) {
        try (XWPFDocument document = new XWPFDocument(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_DOCX,
              MetadataReader.read(document), emit);
          WordParser.parse(document, emit, status);
          EmbeddedObjectParser.parse(document, emit, status);
          finish(status, emit);
        }
      } else if (coreType.contains("spreadsheetml")) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_XLSX,
              MetadataReader.read(workbook), emit);
          SpreadsheetParser.parse(workbook, emit, status);
          EmbeddedObjectParser.parse(workbook, emit, status);
          finish(status, emit);
        }
      } else if (coreType.contains("presentationml")) {
        try (XMLSlideShow show = new XMLSlideShow(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_PPTX,
              MetadataReader.read(show), emit);
          SlideShowParser.parse(show, emit, status);
          EmbeddedObjectParser.parse(show, emit, status);
          finish(status, emit);
        }
      } else {
        throw new UnsupportedFormatException("unsupported OOXML core type: " + coreType);
      }
    }
  }

  private static void parseOle2(String documentId, byte[] bytes, Consumer<ParseEvent> emit)
      throws IOException {
    try (POIFSFileSystem container = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
      DirectoryNode root = container.getRoot();
      if (root.hasEntryCaseInsensitive("WordDocument")) {
        try (HWPFDocument document = new HWPFDocument(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_DOC,
              MetadataReader.read(document.getSummaryInformation()), emit);
          LegacyWordParser.parse(document, emit, status);
          finish(status, emit);
        }
      } else if (root.hasEntryCaseInsensitive("Workbook")) {
        try (HSSFWorkbook workbook = new HSSFWorkbook(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_XLS,
              MetadataReader.read(workbook.getSummaryInformation()), emit);
          SpreadsheetParser.parse(workbook, emit, status);
          finish(status, emit);
        }
      } else if (root.hasEntryCaseInsensitive("PowerPoint Document")) {
        try (HSLFSlideShow show = new HSLFSlideShow(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_PPT,
              MetadataReader.read(show.getSlideShowImpl().getSummaryInformation()), emit);
          SlideShowParser.parse(show, emit, status);
          finish(status, emit);
        }
      } else {
        throw new UnsupportedFormatException("OLE2 container without a known office stream");
      }
    }
  }

  private static ParseStatus.Builder start(
      String documentId, DocumentFormat format, DocumentMetadata metadata,
      Consumer<ParseEvent> emit) {
    emit.accept(
        ParseEvent.newBuilder()
            .setDocumentInfo(
                DocumentInfo.newBuilder()
                    .setDocumentId(documentId)
                    .setFormat(format)
                    .setMetadata(metadata))
            .build());
    return ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
  }

  private static void finish(ParseStatus.Builder status, Consumer<ParseEvent> emit) {
    emit.accept(ParseEvent.newBuilder().setStatus(status).build());
  }

  private static OPCPackage openPackage(byte[] bytes) {
    try {
      return OPCPackage.open(new ByteArrayInputStream(bytes));
    } catch (Exception error) {
      throw new InvalidDocumentException("unreadable OOXML container", error);
    }
  }

  private static String coreContentType(OPCPackage container) {
    var relationships = container.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
    if (relationships.size() == 0) {
      throw new UnsupportedFormatException("OOXML package without a core document");
    }
    PackagePart core = container.getPart(relationships.getRelationship(0));
    if (core == null) throw new UnsupportedFormatException("OOXML core part missing");
    return core.getContentType();
  }
}
