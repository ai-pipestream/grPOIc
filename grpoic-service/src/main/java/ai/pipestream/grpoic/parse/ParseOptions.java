package ai.pipestream.grpoic.parse;

/**
 * What a client asked of one parse, read from the first chunk of its
 * upload.
 *
 * @param sheetBatches stream a large worksheet as several consecutive Sheet
 *     events ({@code Sheet.more_rows}) instead of one. Off by default,
 *     because a consumer that predates batching would read each batch as a
 *     separate sheet.
 */
public record ParseOptions(boolean sheetBatches) {

  /** What a client that sets nothing gets: one Sheet event per worksheet. */
  public static final ParseOptions DEFAULTS = new ParseOptions(false);
}
