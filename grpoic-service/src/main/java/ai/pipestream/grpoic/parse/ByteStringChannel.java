package ai.pipestream.grpoic.parse;

import com.google.protobuf.ByteString;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.ClosedChannelException;
import java.nio.channels.NonWritableChannelException;
import java.nio.channels.SeekableByteChannel;

/**
 * A read-only seekable channel over the upload buffer. A zip reader seeks
 * to the central directory and then to one entry at a time, so a package
 * is read in place: no copy of the document, no entry inflated before
 * something asks for it.
 */
final class ByteStringChannel implements SeekableByteChannel {

  private final ByteString data;
  private long position;
  private boolean open = true;

  ByteStringChannel(ByteString data) {
    this.data = data;
  }

  @Override
  public int read(ByteBuffer destination) throws IOException {
    ensureOpen();
    long size = data.size();
    if (position >= size) return -1;
    int count = (int) Math.min(destination.remaining(), size - position);
    if (count == 0) return 0;
    int start = (int) position;
    data.substring(start, start + count).copyTo(destination);
    position += count;
    return count;
  }

  @Override
  public int write(ByteBuffer source) {
    throw new NonWritableChannelException();
  }

  @Override
  public long position() throws IOException {
    ensureOpen();
    return position;
  }

  @Override
  public SeekableByteChannel position(long newPosition) throws IOException {
    ensureOpen();
    if (newPosition < 0) throw new IllegalArgumentException("negative position " + newPosition);
    position = newPosition;
    return this;
  }

  @Override
  public long size() throws IOException {
    ensureOpen();
    return data.size();
  }

  @Override
  public SeekableByteChannel truncate(long size) {
    throw new NonWritableChannelException();
  }

  @Override
  public boolean isOpen() {
    return open;
  }

  @Override
  public void close() {
    open = false;
  }

  private void ensureOpen() throws ClosedChannelException {
    if (!open) throw new ClosedChannelException();
  }
}
