package ai.pipestream.grpoic.parse;

import ai.pipestream.poi.v1.DocumentMetadata;
import ai.pipestream.poi.v1.MetadataEntry;
import ai.pipestream.poi.v1.MetadataValue;
import com.google.protobuf.Timestamp;
import java.util.Date;
import java.util.Optional;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.PackageProperties;

/**
 * Maps document properties onto the typed metadata contract. Well-known core
 * properties become first-class fields; everything else lands in the tagged
 * tail with the type the document declared, never guessed from string shapes.
 */
public final class MetadataReader {

  private MetadataReader() {}

  /** OOXML core, extended, and custom properties. */
  public static DocumentMetadata read(POIXMLDocument document) {
    DocumentMetadata.Builder metadata = DocumentMetadata.newBuilder();
    POIXMLProperties properties = document.getProperties();
    PackageProperties core = properties.getCoreProperties().getUnderlyingProperties();
    core.getTitleProperty().ifPresent(metadata::setTitle);
    core.getCreatorProperty().ifPresent(metadata::setAuthor);
    core.getLastModifiedByProperty().ifPresent(metadata::setLastModifiedBy);
    core.getCreatedProperty().map(MetadataReader::timestamp).ifPresent(metadata::setCreated);
    core.getModifiedProperty().map(MetadataReader::timestamp).ifPresent(metadata::setModified);
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
      MetadataValue.Builder value = MetadataValue.newBuilder();
      if (property.isSetLpwstr()) {
        value.setStringValue(property.getLpwstr());
      } else if (property.isSetI4()) {
        value.setIntValue(property.getI4());
      } else if (property.isSetR8()) {
        value.setDoubleValue(property.getR8());
      } else if (property.isSetBool()) {
        value.setBoolValue(property.getBool());
      } else if (property.isSetFiletime()) {
        value.setTimestampValue(timestamp(property.getFiletime().getTime()));
      } else {
        continue;
      }
      metadata.addTail(
          MetadataEntry.newBuilder().setKey("custom:" + property.getName()).addValues(value));
    }
    return metadata.build();
  }

  /** HPSF summary stream of the OLE2 legacy formats; null-safe. */
  public static DocumentMetadata read(SummaryInformation summary) {
    DocumentMetadata.Builder metadata = DocumentMetadata.newBuilder();
    if (summary == null) return metadata.build();
    Optional.ofNullable(summary.getTitle()).ifPresent(metadata::setTitle);
    Optional.ofNullable(summary.getAuthor()).ifPresent(metadata::setAuthor);
    Optional.ofNullable(summary.getLastAuthor()).ifPresent(metadata::setLastModifiedBy);
    Optional.ofNullable(summary.getCreateDateTime())
        .map(MetadataReader::timestamp)
        .ifPresent(metadata::setCreated);
    Optional.ofNullable(summary.getLastSaveDateTime())
        .map(MetadataReader::timestamp)
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

  private static Timestamp timestamp(Date date) {
    long millis = date.getTime();
    return Timestamp.newBuilder()
        .setSeconds(Math.floorDiv(millis, 1000L))
        .setNanos((int) (Math.floorMod(millis, 1000L) * 1_000_000L))
        .build();
  }
}
