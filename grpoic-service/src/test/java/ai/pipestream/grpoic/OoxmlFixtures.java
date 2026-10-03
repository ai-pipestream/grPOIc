package ai.pipestream.grpoic;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Hand-written OOXML fixtures: a package POI authors in memory with one part
 * replaced by literal XML. Structures POI's writer never produces (vertical
 * merges, row padding, content-control rows, shared formulas) are spelled
 * out exactly, so every assertion is against XML the test placed.
 */
final class OoxmlFixtures {

  static final String W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  private OoxmlFixtures() {}

  /** A DOCX whose body is the given WordprocessingML (w: prefix bound). */
  static byte[] docxWithBody(String bodyXml) throws Exception {
    byte[] blank;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.write(out);
      blank = out.toByteArray();
    }
    return withPart(blank, "/word/document.xml",
        "<w:document xmlns:w=\"" + W + "\"><w:body>" + bodyXml + "</w:body></w:document>");
  }

  /** The package with one part's bytes replaced by the given XML. */
  static byte[] withPart(byte[] packageBytes, String partName, String xml) throws Exception {
    try (OPCPackage container = OPCPackage.open(new ByteArrayInputStream(packageBytes));
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      PackagePart part = container.getPart(PackagingURIHelper.createPartName(partName));
      try (OutputStream partOut = part.getOutputStream()) {
        partOut.write(xml.getBytes(StandardCharsets.UTF_8));
      }
      container.save(out);
      return out.toByteArray();
    }
  }

  /** A w:tc holding one paragraph of text, with optional cell properties. */
  static String tc(String properties, String text) {
    return "<w:tc>" + (properties.isEmpty() ? "" : "<w:tcPr>" + properties + "</w:tcPr>")
        + "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p></w:tc>";
  }

  static String tc(String text) {
    return tc("", text);
  }
}
