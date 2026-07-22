package ai.pipestream.grpoic.parse;

/** The bytes are not a format this server parses. Maps to UNIMPLEMENTED. */
public class UnsupportedFormatException extends RuntimeException {
  public UnsupportedFormatException(String message) {
    super(message);
  }
}
