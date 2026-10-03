package ai.pipestream.grpoic.parse;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.SAXException;

/**
 * XLSX worksheets read as a stream, one row in memory at a time, with the
 * exact cell semantics of POI's usermodel.
 *
 * <p><b>Why not {@code XSSFWorkbook}.</b> It parses every worksheet into an
 * XMLBeans tree up front: several times the sheet XML in heap, all sheets at
 * once. A workbook with a few million cells needs gigabytes, and the whole
 * tree lives until the parse ends.
 *
 * <p><b>Why not POI's SAX handler.</b> {@code XSSFSheetXMLHandler} hands out
 * display strings only, has no shared-formula support, and formats numbers
 * through a different path than the usermodel does, so the typed cells on
 * the wire would change.
 *
 * <p><b>How.</b> A StAX reader walks the sheet part; each {@code <row>} is
 * parsed on its own into a {@code CTRow} and wrapped in a real
 * {@link XSSFRow}. The row is bound to a scratch workbook that lends it the
 * source's styles, shared strings, date system, sheet names and defined
 * names, so every {@link XSSFCell} computes its type, value, formula
 * (shared and array formulas included) and display string exactly as it
 * would inside a fully loaded workbook. The scratch sheet keeps only what
 * later rows refer back to: shared-formula masters and array-formula
 * anchors. Shared strings are held by the read-only SAX table, a plain
 * string list rather than an XMLBeans tree.
 *
 * <p><b>Same sheets, same order.</b> Sheets are taken from the workbook
 * part in workbook order, and only the kinds {@code XSSFWorkbook} loads
 * (worksheets and chart sheets) count, so indices match the usermodel.
 */
final class XlsxSheets {

  private static final XmlOptions ROW_OPTIONS =
      new XmlOptions(POIXMLTypeLoader.DEFAULT_XML_OPTIONS);
  private static final String MAIN = XSSFRelation.NS_SPREADSHEETML;

  static {
    // Parse a lone <row> element as the CTRow itself, not as a document
    // holding one.
    ROW_OPTIONS.setLoadReplaceDocumentElement(null);
  }

  /** One worksheet: its name, visibility, part, and the scratch sheet its rows bind to. */
  record Worksheet(String name, boolean hidden, PackagePart part, XSSFSheet scratch) {}

  private final List<Worksheet> sheets;

  private XlsxSheets(List<Worksheet> sheets) {
    this.sheets = sheets;
  }

  static XlsxSheets open(OPCPackage container)
      throws IOException, XmlException, SAXException, OpenXML4JException {
    PackageRelationship coreRelationship =
        container.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT).getRelationship(0);
    PackagePart workbookPart = container.getPart(coreRelationship);
    CTWorkbook workbook;
    try (InputStream in = workbookPart.getInputStream()) {
      workbook = WorkbookDocument.Factory.parse(in, POIXMLTypeLoader.DEFAULT_XML_OPTIONS)
          .getWorkbook();
    }
    XSSFReader reader = new XSSFReader(container);
    ScratchWorkbook scratch = new ScratchWorkbook(
        reader.getStylesTable(),
        new StreamedStrings(new ReadOnlySharedStringsTable(container, false)),
        workbook.isSetWorkbookPr() && workbook.getWorkbookPr().getDate1904());

    List<Worksheet> sheets = new ArrayList<>();
    if (workbook.getSheets() != null) {
      for (CTSheet sheet : workbook.getSheets().getSheetArray()) {
        PackageRelationship relationship =
            sheet.getId() == null ? null : workbookPart.getRelationship(sheet.getId());
        if (relationship == null || !isLoadedSheet(relationship)) continue;
        PackagePart part = workbookPart.getRelatedPart(relationship);
        if (part == null) continue;
        boolean hidden = sheet.isSetState() && sheet.getState() != STSheetState.VISIBLE;
        sheets.add(new Worksheet(
            sheet.getName(), hidden, part, scratch.addSheet(sheets.size())));
      }
    }
    if (workbook.isSetDefinedNames()) {
      scratch.mirrorNames(workbook.getDefinedNames().getDefinedNameArray());
    }
    return new XlsxSheets(sheets);
  }

  List<Worksheet> sheets() {
    return sheets;
  }

  /** The OLE and package embeddings of every sheet, as XSSFWorkbook lists them. */
  List<PackagePart> embeddedParts() throws InvalidFormatException {
    List<PackagePart> parts = new ArrayList<>();
    for (Worksheet sheet : sheets) {
      for (String type : List.of(XSSFRelation.OLEEMBEDDINGS.getRelation(),
                                 XSSFRelation.PACKEMBEDDINGS.getRelation())) {
        for (PackageRelationship relationship : sheet.part().getRelationshipsByType(type)) {
          parts.add(sheet.part().getRelatedPart(relationship));
        }
      }
    }
    return parts;
  }

  /**
   * The sheet's rows in stored order, each a real XSSFRow. Single use; the
   * part stream closes when the rows run out or {@link Rows#close()} runs.
   */
  static Rows rows(Worksheet sheet) throws IOException {
    InputStream in = sheet.part().getInputStream();
    try {
      return new Rows(
          XMLHelper.newXMLInputFactory().createXMLStreamReader(in), in, sheet.scratch());
    } catch (XMLStreamException error) {
      in.close();
      throw new InvalidDocumentException("unreadable worksheet: " + error.getMessage(), error);
    }
  }

  private static boolean isLoadedSheet(PackageRelationship relationship) {
    String type = relationship.getRelationshipType();
    return XSSFRelation.WORKSHEET.getRelation().equals(type)
        || XSSFRelation.CHARTSHEET.getRelation().equals(type);
  }

  /**
   * Streaming row iterator over one sheet part. The merged-cell references
   * the sheet declares after its rows are collected on the way to the end.
   */
  static final class Rows implements Iterator<Row>, AutoCloseable {
    private final XMLStreamReader reader;
    private final InputStream in;
    private final XSSFSheet scratch;
    private final List<String> mergedReferences = new ArrayList<>();
    private boolean inSheetData;
    private boolean inMergeCells;
    private long lastRow;
    private Row next;
    private boolean done;

    private Rows(XMLStreamReader reader, InputStream in, XSSFSheet scratch) {
      this.reader = reader;
      this.in = in;
      this.scratch = scratch;
    }

    @Override
    public boolean hasNext() {
      if (next == null && !done) advance();
      return next != null;
    }

    @Override
    public Row next() {
      if (!hasNext()) throw new NoSuchElementException();
      Row row = next;
      next = null;
      return row;
    }

    /**
     * The {@code ref} of every merged-cell range the sheet declares, valid
     * once the rows have run out.
     */
    List<String> mergedReferences() {
      return mergedReferences;
    }

    private void advance() {
      try {
        while (reader.hasNext()) {
          int event = reader.next();
          if (event == XMLStreamConstants.START_ELEMENT && MAIN.equals(reader.getNamespaceURI())) {
            String element = reader.getLocalName();
            if (element.equals("sheetData")) {
              inSheetData = true;
            } else if (inSheetData && element.equals("row")) {
              next = row(CTRow.Factory.parse(reader, ROW_OPTIONS));
              return;
            } else if (element.equals("mergeCells")) {
              inMergeCells = true;
            } else if (inMergeCells && element.equals("mergeCell")) {
              String reference = reader.getAttributeValue(null, "ref");
              // One past the cap is kept, so the converter can tell it was hit.
              if (reference != null
                  && mergedReferences.size() <= SpreadsheetParser.MAX_MERGED_REGIONS) {
                mergedReferences.add(reference);
              }
            }
          } else if (event == XMLStreamConstants.END_ELEMENT
              && MAIN.equals(reader.getNamespaceURI())) {
            String element = reader.getLocalName();
            if (element.equals("sheetData")) inSheetData = false;
            if (element.equals("mergeCells")) inMergeCells = false;
          }
        }
        close();
      } catch (XMLStreamException | XmlException error) {
        close();
        throw new InvalidDocumentException("malformed worksheet: " + error.getMessage(), error);
      }
    }

    private Row row(CTRow stored) {
      // A writer may omit row numbers; like XSSFSheet, take the next one.
      if (!stored.isSetR()) stored.setR(lastRow + 1);
      lastRow = stored.getR();
      XSSFRow row = new StreamedRow(stored, scratch);
      for (Cell cell : row) {
        CTCell source = ((XSSFCell) cell).getCTCell();
        if (source.isSetF() && source.getF().getT() == STCellFormulaType.ARRAY
            && source.getF().isSetRef()) {
          // The cells an array formula covers ask the sheet for its anchor.
          XSSFRow anchors = scratch.getRow(cell.getRowIndex());
          if (anchors == null) anchors = scratch.createRow(cell.getRowIndex());
          anchors.createCell(cell.getColumnIndex()).getCTCell().set(source.copy());
        }
      }
      return row;
    }

    @Override
    public void close() {
      if (done) return;
      done = true;
      try {
        reader.close();
        in.close();
      } catch (XMLStreamException | IOException ignored) {
        // in-memory sources; nothing to recover
      }
    }
  }

  /** XSSFRow's constructor is protected; this binds a parsed row to a sheet. */
  private static final class StreamedRow extends XSSFRow {
    StreamedRow(CTRow row, XSSFSheet sheet) {
      super(row, sheet);
    }
  }

  /**
   * An empty workbook that answers for the source: its styles, shared
   * strings and date system, plus one sheet per source sheet and the
   * source's defined names, which shared formulas resolve against.
   */
  private static final class ScratchWorkbook extends XSSFWorkbook {
    private final StylesTable styles;
    private final SharedStringsTable strings;
    private final boolean date1904;

    ScratchWorkbook(StylesTable styles, SharedStringsTable strings, boolean date1904) {
      this.styles = styles;
      this.strings = strings;
      this.date1904 = date1904;
    }

    // Fields are unset while XSSFWorkbook's own constructor runs; it gets
    // its defaults then. A package without a styles part keeps them.
    @Override
    public StylesTable getStylesSource() {
      return styles != null ? styles : super.getStylesSource();
    }

    @Override
    public SharedStringsTable getSharedStringSource() {
      return strings != null ? strings : super.getSharedStringSource();
    }

    @Override
    public boolean isDate1904() {
      return date1904;
    }

    /** A sheet for each source sheet, by position; names here are internal only. */
    XSSFSheet addSheet(int index) {
      return createSheet("sheet" + (index + 1));
    }

    /**
     * The source's defined names, so shared formulas that use them render.
     * Names are created first and given their formulas second, because a
     * name may refer to one defined after it. A name POI cannot recreate is
     * left out; only a formula that uses it is affected.
     */
    void mirrorNames(CTDefinedName[] names) {
      List<XSSFName> created = new ArrayList<>();
      for (CTDefinedName name : names) {
        try {
          XSSFName copy = createName();
          if (name.isSetLocalSheetId()) copy.setSheetIndex((int) name.getLocalSheetId());
          copy.setNameName(name.getName());
          if (name.getFunction()) copy.setFunction(true);
          created.add(copy);
        } catch (RuntimeException unusable) {
          created.add(null);
        }
      }
      for (int index = 0; index < names.length; index++) {
        if (created.get(index) == null) continue;
        try {
          created.get(index).setRefersToFormula(names[index].getStringValue());
        } catch (RuntimeException unparseable) {
          // the name exists without a formula
        }
      }
    }
  }

  /**
   * Shared strings from the read-only SAX table, a plain string list,
   * through the usermodel's type. XSSFCell asks its workbook for this type,
   * and the table it normally gets holds an XMLBeans tree per string.
   */
  private static final class StreamedStrings extends SharedStringsTable {
    private final ReadOnlySharedStringsTable source;

    StreamedStrings(ReadOnlySharedStringsTable source) {
      this.source = source;
    }

    @Override
    public RichTextString getItemAt(int index) {
      return source.getItemAt(index);
    }

    @Override
    public int getCount() {
      return source.getCount();
    }

    @Override
    public int getUniqueCount() {
      return source.getUniqueCount();
    }
  }
}
