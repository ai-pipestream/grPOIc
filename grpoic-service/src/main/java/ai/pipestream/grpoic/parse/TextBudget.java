package ai.pipestream.grpoic.parse;

import org.apache.poi.openxml4j.util.ZipSecureFile;

/**
 * Characters of text a spreadsheet may still yield, under the limit
 * {@link PoiLimits} installs. Spending past it ends the parse with
 * {@link DocumentTooLargeException}, so a small file cannot expand into an
 * unbounded amount of text in heap or on the wire.
 */
final class TextBudget {

  private final String what;
  private final long limit;
  private long spent;

  /** A budget at the installed limit; {@code what} names it in the failure. */
  TextBudget(String what) {
    this.what = what;
    this.limit = ZipSecureFile.getMaxTextSize();
  }

  void spend(long characters) {
    spent += characters;
    if (spent > limit) {
      throw new DocumentTooLargeException(what + " exceeds " + limit + " characters");
    }
  }
}
