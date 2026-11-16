using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void PooledReadBufferStreamReturnsFileBytesForMixedSeeksAndReadSizes() {
        var random = new Random(17);
        byte[] content = new byte[20_000];
        random.NextBytes(content);
        string path = Path.Combine(_directoryWithFiles, "PooledReadBufferStream.bin");
        File.WriteAllBytes(path, content);

        using (var stream = ExcelPooledReadBufferStream.OpenFile(path, bufferSize: 1024)) {
            Assert.Equal(content.Length, stream.Length);
            for (int step = 0; step < 400; step++) {
                long position = random.Next(content.Length + 10);
                stream.Seek(position, SeekOrigin.Begin);
                if (step % 7 == 0) {
                    Assert.Equal(position < content.Length ? content[position] : -1, stream.ReadByte());
                    continue;
                }

                // Sizes below, at and above the buffer cover buffered copies, refills and direct reads.
                int requested = step % 3 == 0 ? random.Next(1024, 3000) : random.Next(1, 300);
                byte[] buffer = new byte[requested];
                int total = 0;
                int read;
                while (total < requested && (read = stream.Read(buffer, total, requested - total)) > 0) total += read;

                int expected = (int)Math.Max(0, Math.Min(requested, content.Length - position));
                Assert.Equal(expected, total);
                byte[] expectedBytes = new byte[expected];
                Array.Copy(content, Math.Min(position, content.Length), expectedBytes, 0, expected);
                Assert.Equal(expectedBytes, buffer.Take(total).ToArray());
                Assert.Equal(position + total, stream.Position);
            }

            stream.Dispose();
            Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
        }
    }
}
