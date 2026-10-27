package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.DocumentMetadata;
import ai.pipestream.poi.v1.MetadataEntry;
import ai.pipestream.poi.v1.MetadataValue;
import java.util.Calendar;
import java.util.Optional;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.PackageProperties;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

/**
 * Maps document properties onto the typed metadata contract. Well-known core
 * properties become first-class fields; everything else lands in the tagged
 * tail with the type the document declared, never guessed from string shapes.
 */
public final class MetadataReader {

  private MetadataReader() {}

  /** OOXML core, extended, and custom properties. */
  public static DocumentMetadata read(POIXMLDocument document) {
    return read(document.getProperties());
  }

  /** OOXML core, extended, and custom properties read straight from a package. */
  public static DocumentMetadata read(POIXMLProperties properties) {
    DocumentMetadata.Builder metadata = DocumentMetadata.newBuilder();
    PackageProperties core = properties.getCoreProperties().getUnderlyingProperties();
    core.getTitleProperty().ifPresent(metadata::setTitle);
    core.getCreatorProperty().ifPresent(metadata::setAuthor);
    core.getLastModifiedByProperty().ifPresent(metadata::setLastModifiedBy);
    core.getCreatedProperty().map(ProtoTimestamps::fromDate).ifPresent(metadata::setCreated);
    core.getModifiedProperty().map(ProtoTimestamps::fromDate).ifPresent(metadata::setModified);
    core.getSubjectProperty().ifPresent(value -> addString(metadata, "subject", value));
    core.getKeywordsProperty().ifPresent(value -> addString(metadata, "keywords", value));
    core.getDescriptionProperty().ifPresent(value -> addString(metadata, "description", value));
    core.getCategoryProperty().ifPresent(value -> addString(metadata, "category", value));
    core.getRevisionProperty().ifPresent(value -> addString(metadata, "revision", value));
    core.getLanguageProperty().ifPresent(value -> addString(metadata, "language", value));

    var extended = properties.getExtendedProperties().getUnderlyingProperties();
    if (extended.getApplication() != null) {
      addString(metadata, "application", extended.getApplication());
    }
    if (extended.isSetAppVersion()) {
      addString(metadata, "application_version", extended.getAppVersion());
    }
    if (extended.isSetCompany() && !extended.getCompany().isEmpty()) {
      addString(metadata, "company", extended.getCompany());
    }
    if (extended.isSetPages()) addInt(metadata, "pages", extended.getPages());
    if (extended.isSetWords()) addInt(metadata, "words", extended.getWords());
    if (extended.isSetCharacters()) addInt(metadata, "characters", extended.getCharacters());

    for (var property : properties.getCustomProperties().getUnderlyingProperties().getPropertyList()) {
      MetadataValue value = customValue(property);
      if (value == null) continue;
      metadata.addTail(
          MetadataEntry.newBuilder().setKey("custom:" + property.getName()).addValues(value));
    }
    return metadata.build();
  }

  /**
   * A custom property's value in the type it declares, or null for a type
   * not carried on the wire or a value the part marks nil (no value).
   */
  private static MetadataValue customValue(CTProperty property) {
    MetadataValue.Builder value = MetadataValue.newBuilder();
    if (property.isSetLpwstr()) {
      String text = property.getLpwstr();
      return text == null ? null : value.setStringValue(text).build();
    }
    if (property.isSetFiletime()) {
      Calendar time = property.getFiletime();
      if (time == null) return null;
      return value.setTimestampValue(ProtoTimestamps.fromDate(time.getTime())).build();
    }
    if (property.isSetI4()) return value.setIntValue(property.getI4()).build();
    if (property.isSetR8()) return value.setDoubleValue(property.getR8()).build();
    if (property.isSetBool()) return value.setBoolValue(property.getBool()).build();
    return null;
  }

  /** HPSF summary stream of the OLE2 legacy formats; null-safe. */
  public static DocumentMetadata read(SummaryInformation summary) {
    DocumentMetadata.Builder metadata = DocumentMetadata.newBuilder();
    if (summary == null) return metadata.build();
    Optional.ofNullable(summary.getTitle()).ifPresent(metadata::setTitle);
    Optional.ofNullable(summary.getAuthor()).ifPresent(metadata::setAuthor);
    Optional.ofNullable(summary.getLastAuthor()).ifPresent(metadata::setLastModifiedBy);
    Optional.ofNullable(summary.getCreateDateTime())
        .map(ProtoTimestamps::fromDate)
        .ifPresent(metadata::setCreated);
    Optional.ofNullable(summary.getLastSaveDateTime())
        .map(ProtoTimestamps::fromDate)
        .ifPresent(metadata::setModified);
    Optional.ofNullable(summary.getApplicationName())
        .filter(name -> !name.isEmpty())
        .ifPresent(name -> addString(metadata, "application", name));
    Optional.ofNullable(summary.getKeywords())
        .filter(keywords -> !keywords.isEmpty())
        .ifPresent(keywords -> addString(metadata, "keywords", keywords));
    Optional.ofNullable(summary.getComments())
        .filter(comments -> !comments.isEmpty())
        .ifPresent(comments -> addString(metadata, "description", comments));
    return metadata.build();
  }

  private static void addString(DocumentMetadata.Builder metadata, String key, String value) {
    metadata.addTail(
        MetadataEntry.newBuilder()
            .setKey(key)
            .addValues(MetadataValue.newBuilder().setStringValue(value)));
  }

  private static void addInt(DocumentMetadata.Builder metadata, String key, long value) {
    metadata.addTail(
        MetadataEntry.newBuilder()
            .setKey(key)
            .addValues(MetadataValue.newBuilder().setIntValue(value)));
  }
}
