package ai.pipestream.grpoic.parse;

import com.google.protobuf.Timestamp;
import java.util.Date;

/** Conversion from POI's legacy {@link Date} values to protobuf timestamps. */
final class ProtoTimestamps {

  private ProtoTimestamps() {}

  static Timestamp fromDate(Date date) {
    long millis = date.getTime();
    return Timestamp.newBuilder()
        .setSeconds(Math.floorDiv(millis, 1000L))
        .setNanos((int) (Math.floorMod(millis, 1000L) * 1_000_000L))
        .build();
  }
}
