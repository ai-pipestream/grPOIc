package ai.pipestream.grpoic.parse;

/**
 * The document is encrypted and cannot be read without its password.
 * Maps to FAILED_PRECONDITION: the bytes are fine, the server lacks a key.
 */
public class ProtectedDocumentException extends RuntimeException {
  public ProtectedDocumentException(String message, Throwable cause) {
    super(message, cause);
  }
}
