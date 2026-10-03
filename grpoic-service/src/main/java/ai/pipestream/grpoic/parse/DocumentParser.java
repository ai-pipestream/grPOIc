package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.DocumentFormat;
import ai.pipestream.poi.v1.DocumentInfo;
import ai.pipestream.poi.v1.DocumentMetadata;
import ai.pipestream.poi.v1.ParseEvent;
import ai.pipestream.poi.v1.ParseStatus;
import com.google.protobuf.ByteString;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.OldFileFormatException;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXException;

/**
 * Detects the format from the bytes (the advisory content type is never
 * trusted) and streams typed events: DocumentInfo first, content blocks in
 * document order, ParseStatus last. Entirely in memory, read in place from
 * the upload buffer; nothing is written to disk and no process is executed.
 * Per-format extraction lives in the sibling parser classes; this class only
 * detects, dispatches, and frames the stream.
 */
public final class DocumentParser {

  // Main-part content types per family: documents, templates, and their
  // macro-enabled variants. Macros are never run; POI has no VBA engine and
  // the vbaProject part is never read.
  private static final Set<String> WORD_TYPES = contentTypes(
      XWPFRelation.DOCUMENT.getContentType(),
      XWPFRelation.TEMPLATE.getContentType(),
      XWPFRelation.MACRO_DOCUMENT.getContentType(),
      XWPFRelation.MACRO_TEMPLATE_DOCUMENT.getContentType());
  private static final Set<String> SPREADSHEET_TYPES = contentTypes(
      XSSFRelation.WORKBOOK.getContentType(),
      XSSFRelation.TEMPLATE_WORKBOOK.getContentType(),
      XSSFRelation.MACROS_WORKBOOK.getContentType(),
      XSSFRelation.MACRO_TEMPLATE_WORKBOOK.getContentType(),
      XSSFRelation.MACRO_ADDIN_WORKBOOK.getContentType());
  private static final Set<String> PRESENTATION_TYPES = contentTypes(
      XSLFRelation.MAIN.getContentType(),
      XSLFRelation.PRESENTATIONML.getContentType(),
      XSLFRelation.PRESENTATIONML_TEMPLATE.getContentType(),
      XSLFRelation.PRESENTATION_MACRO.getContentType(),
      XSLFRelation.MACRO.getContentType(),
      XSLFRelation.MACRO_TEMPLATE.getContentType());

  private DocumentParser() {}

  /**
   * Parses one document. Failures leave as one of three typed exceptions the
   * service maps to a status: {@link UnsupportedFormatException} (not a format
   * parsed here), {@link InvalidDocumentException} (the bytes claim a format
   * but are broken), {@link ProtectedDocumentException} (encrypted). Anything
   * else escaping is a server fault, including whatever {@code emit} throws.
   */
  public static void parse(String documentId, ByteString data, ParseOptions options,
                           Consumer<ParseEvent> emit) {
    FileMagic magic = FileMagic.valueOf(data.substring(0, Math.min(data.size(), 64)).toByteArray());
    try {
      switch (magic) {
        case OOXML -> parseOoxml(documentId, data, options, emit);
        case OLE2 -> parseOle2(documentId, data, options, emit);
        default -> throw new UnsupportedFormatException(
            "not an office document (magic: " + magic + ")");
      }
    } catch (UnsupportedFormatException | InvalidDocumentException
             | ProtectedDocumentException classified) {
      throw classified;
    } catch (EncryptedDocumentException encrypted) {
      throw new ProtectedDocumentException(
          "encrypted document: " + DocumentFaults.describe(encrypted), encrypted);
    } catch (OldFileFormatException old) {
      throw new UnsupportedFormatException(
          "pre-97 office format: " + DocumentFaults.describe(old));
    } catch (IOException | OpenXML4JException | XmlException | SAXException error) {
      throw new InvalidDocumentException("unreadable document: " + error.getMessage(), error);
    } catch (RuntimeException error) {
      if (!DocumentFaults.fromDocument(error)) throw error;
      throw new InvalidDocumentException(
          "malformed document: " + DocumentFaults.describe(error), error);
    }
  }

  private static void parseOoxml(String documentId, ByteString data, ParseOptions options,
                                 Consumer<ParseEvent> emit)
      throws IOException, OpenXML4JException, XmlException, SAXException {
    try (OPCPackage container = openPackage(data)) {
      String coreType = coreContentType(container);
      String normalized = coreType.toLowerCase(Locale.ROOT);
      if (WORD_TYPES.contains(normalized)) {
        try (XWPFDocument document = new XWPFDocument(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_DOCX,
              () -> MetadataReader.read(document), emit);
          WordParser.parse(document, emit, status);
          EmbeddedObjectParser.parse(
              () -> EmbeddedObjectParser.embeddings(List.of(document.getPackagePart())),
              emit, status);
          finish(status, emit);
        }
      } else if (SPREADSHEET_TYPES.contains(normalized)) {
        XlsxSheets workbook = XlsxSheets.open(container);
        ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_XLSX,
            () -> MetadataReader.read(properties(container)), emit);
        SpreadsheetParser.parse(workbook, options, emit, status);
        EmbeddedObjectParser.parse(workbook::embeddedParts, emit, status);
        finish(status, emit);
      } else if (PRESENTATION_TYPES.contains(normalized)) {
        try (XMLSlideShow show = new XMLSlideShow(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_PPTX,
              () -> MetadataReader.read(show), emit);
          SlideShowParser.parse(show, emit, status);
          EmbeddedObjectParser.parse(show, emit, status);
          finish(status, emit);
        }
      } else {
        throw new UnsupportedFormatException("unsupported OOXML core type: " + coreType);
      }
    }
  }

  private static void parseOle2(String documentId, ByteString data, ParseOptions options,
                                Consumer<ParseEvent> emit) throws IOException {
    try (POIFSFileSystem container = new POIFSFileSystem(data.newInput())) {
      DirectoryNode root = container.getRoot();
      if (root.hasEntryCaseInsensitive("EncryptedPackage")) {
        // An encrypted DOCX/XLSX/PPTX travels as an OLE2 container holding
        // the encrypted zip; without the password there is nothing to read.
        throw new ProtectedDocumentException("encrypted OOXML package", null);
      } else if (root.hasEntryCaseInsensitive("WordDocument")) {
        try (HWPFDocument document = new HWPFDocument(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_DOC,
              () -> MetadataReader.read(document.getSummaryInformation()), emit);
          LegacyWordParser.parse(document, emit, status);
          finish(status, emit);
        }
      } else if (root.hasEntryCaseInsensitive("Workbook")) {
        try (HSSFWorkbook workbook = new HSSFWorkbook(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_XLS,
              () -> MetadataReader.read(workbook.getSummaryInformation()), emit);
          SpreadsheetParser.parse(workbook, options, emit, status);
          finish(status, emit);
        }
      } else if (root.hasEntryCaseInsensitive("PowerPoint Document")) {
        try (HSLFSlideShow show = new HSLFSlideShow(container)) {
          ParseStatus.Builder status = start(documentId, DocumentFormat.DOCUMENT_FORMAT_LEGACY_PPT,
              () -> MetadataReader.read(show.getSlideShowImpl().getSummaryInformation()), emit);
          SlideShowParser.parse(show, emit, status);
          finish(status, emit);
        }
      } else {
        throw new UnsupportedFormatException("OLE2 container without a known office stream");
      }
    }
  }

  /**
   * Emits DocumentInfo and opens the status. Broken property parts cost the
   * metadata, not the document.
   */
  private static ParseStatus.Builder start(
      String documentId, DocumentFormat format, Supplier<DocumentMetadata> metadata,
      Consumer<ParseEvent> emit) {
    ParseStatus.Builder status = ParseStatus.newBuilder().setState(ParseStatus.State.STATE_OK);
    DocumentMetadata read;
    try {
      read = metadata.get();
    } catch (RuntimeException error) {
      DocumentFaults.skip(status, "document metadata", error);
      read = DocumentMetadata.getDefaultInstance();
    }
    emit.accept(
        ParseEvent.newBuilder()
            .setDocumentInfo(
                DocumentInfo.newBuilder()
                    .setDocumentId(documentId)
                    .setFormat(format)
                    .setMetadata(read))
            .build());
    return status;
  }

  private static void finish(ParseStatus.Builder status, Consumer<ParseEvent> emit) {
    emit.accept(ParseEvent.newBuilder().setStatus(status).build());
  }

  private static OPCPackage openPackage(ByteString data) {
    try {
      return InMemoryPackages.open(data);
    } catch (Exception error) {
      throw new InvalidDocumentException("unreadable OOXML container", error);
    }
  }

  /** Package properties for metadata; a broken property part is a document fault. */
  private static POIXMLProperties properties(OPCPackage container) {
    try {
      return new POIXMLProperties(container);
    } catch (IOException | OpenXML4JException | XmlException error) {
      throw new POIXMLException(error);
    }
  }

  private static Set<String> contentTypes(String... types) {
    return Arrays.stream(types)
        .map(type -> type.toLowerCase(Locale.ROOT))
        .collect(Collectors.toUnmodifiableSet());
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
