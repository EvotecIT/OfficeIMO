using System.IO.Compression;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // Region edges: the first region holds 160 KiB, each later one 128 KiB more.
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(163_839)]
        [InlineData(163_840)]
        [InlineData(163_841)]
        [InlineData(294_912)]
        [InlineData(3_000_017)]
        public void DeflateEncoder_RoundTripsWithTheInputCrcAndTheSameBytesWhereverRegionsRun(int length) {
            byte[] input = CreateDeflateInput(length);
            (byte[] inline, uint inlineCrc) = Deflate(input, background: false);
            (byte[] background, uint backgroundCrc) = Deflate(input, background: true);

            Assert.Equal(inline, background);
            // The pooled match tables still hold another part's positions; the bytes must not depend on them.
            byte[] other = CreateDeflateInput(length + 4_099);
            Array.Reverse(other);
            Deflate(other, background: false);
            Assert.Equal(inline, Deflate(input, background: false).Compressed);
            uint expectedCrc = BitwiseCrc32(input);
            Assert.Equal(expectedCrc, inlineCrc);
            Assert.Equal(expectedCrc, backgroundCrc);
            Assert.Equal(input, Inflate(inline));
        }

        [Fact]
        public void DeflateEncoder_SurfacesADestinationFailureAndStillReleasesRunningRegions() {
            byte[] input = CreateDeflateInput(2_000_000);
            var destination = new FailingDeflateDestination();
            var encoder = new ExcelDeflateEncoder(destination, background: true);
            IOException error = Assert.Throws<IOException>(() => {
                encoder.Write(input);
                encoder.Finish();
            });
            Assert.Same(destination.Failure, error);
            encoder.Dispose();

            // The next part still compresses correctly after the failed one gave back its buffers.
            (byte[] compressed, _) = Deflate(input, background: true);
            Assert.Equal(input, Inflate(compressed));
        }

        [Fact]
        public void DeflateEncoder_DisposeWithoutFinishWaitsForQueuedRegions() {
            byte[] input = CreateDeflateInput(1_500_000);
            var destination = new MemoryStream();
            using (var encoder = new ExcelDeflateEncoder(destination, background: true)) encoder.Write(input);

            (byte[] compressed, _) = Deflate(input, background: true);
            Assert.Equal(input, Inflate(compressed));
        }

        private static (byte[] Compressed, uint Crc) Deflate(byte[] input, bool background) {
            var output = new MemoryStream();
            using var encoder = new ExcelDeflateEncoder(output, background);
            // Uneven writes cross region edges in the middle of a call.
            for (int offset = 0; offset < input.Length; offset += 70_001)
                encoder.Write(new ReadOnlySpan<byte>(input, offset, Math.Min(70_001, input.Length - offset)));
            encoder.Finish();
            Assert.Equal(input.Length, encoder.Size);
            return (output.ToArray(), encoder.Crc);
        }

        // An independent bit-at-a-time CRC-32 (ISO-HDLC), the value ZIP stores.
        private static uint BitwiseCrc32(byte[] bytes) {
            uint crc = 0xFFFF_FFFFU;
            foreach (byte value in bytes) {
                crc ^= value;
                for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB8_8320U : crc >> 1;
            }
            return ~crc;
        }

        private static byte[] Inflate(byte[] compressed) {
            using var inflater = new DeflateStream(new MemoryStream(compressed), CompressionMode.Decompress);
            var result = new MemoryStream();
            inflater.CopyTo(result);
            return result.ToArray();
        }

        // Repetitive rows with changing numbers, like worksheet XML, followed by bytes that do not compress.
        private static byte[] CreateDeflateInput(int length) {
            var input = new byte[length];
            var random = new Random(1951);
            int position = 0, row = 0;
            while (position < length * 3 / 4) {
                byte[] line = System.Text.Encoding.ASCII.GetBytes($"<row r=\"{row}\"><c r=\"A{row}\"><v>{row * 7919 % 100_003}</v></c></row>");
                int count = Math.Min(line.Length, length - position);
                Array.Copy(line, 0, input, position, count);
                position += count;
                row++;
            }
            random.NextBytes(input.AsSpan(position));
            return input;
        }

        private sealed class FailingDeflateDestination : MemoryStream {
            internal IOException Failure { get; } = new("Destination write failure.");
            public override void Write(byte[] buffer, int offset, int count) => throw Failure;
        }
    }
}
