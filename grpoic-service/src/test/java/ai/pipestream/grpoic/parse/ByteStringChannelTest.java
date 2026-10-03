package ai.pipestream.grpoic.parse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.google.protobuf.ByteString;
import java.nio.ByteBuffer;
import java.nio.channels.ClosedChannelException;
import java.nio.channels.NonWritableChannelException;
import org.junit.jupiter.api.Test;

/** Seek and read semantics over a chunked (rope) upload buffer. */
class ByteStringChannelTest {

  private static final ByteString ROPE = ByteString.copyFromUtf8("hello ")
      .concat(ByteString.copyFromUtf8("chunked "))
      .concat(ByteString.copyFromUtf8("world"));

  @Test
  void readsAcrossChunkBoundariesFromAnyPosition() throws Exception {
    try (ByteStringChannel channel = new ByteStringChannel(ROPE)) {
      assertThat(channel.size()).isEqualTo(19);
      channel.position(4);
      ByteBuffer buffer = ByteBuffer.allocate(10);
      assertThat(channel.read(buffer)).isEqualTo(10);
      assertThat(new String(buffer.array())).isEqualTo("o chunked ");
      assertThat(channel.position()).isEqualTo(14);
      buffer.clear();
      assertThat(channel.read(buffer)).as("short read at the end").isEqualTo(5);
      assertThat(channel.read(buffer)).as("end of data").isEqualTo(-1);
      channel.position(100);
      assertThat(channel.read(ByteBuffer.allocate(1))).isEqualTo(-1);
    }
  }

  @Test
  void isReadOnlyAndHonoursClose() throws Exception {
    ByteStringChannel channel = new ByteStringChannel(ROPE);
    assertThatThrownBy(() -> channel.write(ByteBuffer.allocate(1)))
        .isInstanceOf(NonWritableChannelException.class);
    assertThatThrownBy(() -> channel.truncate(1)).isInstanceOf(NonWritableChannelException.class);
    channel.close();
    assertThat(channel.isOpen()).isFalse();
    assertThatThrownBy(() -> channel.read(ByteBuffer.allocate(1)))
        .isInstanceOf(ClosedChannelException.class);
  }
}
