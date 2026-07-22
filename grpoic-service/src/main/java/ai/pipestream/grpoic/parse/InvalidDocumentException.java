package ai.pipestream.grpoic.parse;

/** The bytes claim a supported format but cannot be read. Maps to INVALID_ARGUMENT. */
public class InvalidDocumentException extends RuntimeException {
  public InvalidDocumentException(String message, Throwable cause) {
    super(message, cause);
  }
}
