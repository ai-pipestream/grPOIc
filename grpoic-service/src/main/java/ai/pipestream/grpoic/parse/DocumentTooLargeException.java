package ai.pipestream.grpoic.parse;

/**
 * The document is readable but would yield more than the server's limits
 * allow. Maps to RESOURCE_EXHAUSTED.
 */
public class DocumentTooLargeException extends RuntimeException {
  public DocumentTooLargeException(String message) {
    super(message);
  }
}
