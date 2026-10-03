package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.google.protobuf.ByteString;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

/**
 * Entry names get the checks POI's own zip file applies before the package
 * is opened in place: two entries one reader could resolve to different
 * bytes than another are refused, and so is an entry with no name.
 */
class InMemoryPackagesTest {

  /** A real DOCX with {@code extra} written as an additional entry after word/document.xml. */
  private static ByteString docxWithEntry(String extra) throws Exception {
    byte[] docx;
    try (XWPFDocument document = new XWPFDocument();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      document.createParagraph().createRun().setText("body");
      document.write(out);
      docx = out.toByteArray();
    }
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    try (ZipInputStream in = new ZipInputStream(new ByteArrayInputStream(docx));
         ZipOutputStream zip = new ZipOutputStream(out)) {
      for (ZipEntry entry = in.getNextEntry(); entry != null; entry = in.getNextEntry()) {
        byte[] content = in.readAllBytes();
        zip.putNextEntry(new ZipEntry(entry.getName()));
        zip.write(content);
        if (entry.getName().equals("word/document.xml")) {
          zip.putNextEntry(new ZipEntry(extra));
          zip.write(content);
        }
      }
    }
    return ByteString.copyFrom(out.toByteArray());
  }

  @Test
  void entryNamesDifferingOnlyByCaseAreRefused() throws Exception {
    ByteString data = docxWithEntry("WORD/Document.xml");
    assertThatThrownBy(() -> InMemoryPackages.open(data))
        .isInstanceOf(IOException.class)
        .hasMessage("more than one zip entry named WORD/Document.xml");
  }

  @Test
  void entryWithAnEmptyNameIsRefused() throws Exception {
    ByteString data = docxWithEntry("");
    assertThatThrownBy(() -> InMemoryPackages.open(data))
        .isInstanceOf(IOException.class)
        .hasMessage("zip entry with an empty name");
  }
}
