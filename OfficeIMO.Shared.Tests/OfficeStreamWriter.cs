using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class OfficeStreamWriterTests {
        [Fact]
        public void WriteAllBytesTruncatesAndRewindsSeekableDestination() {
            using var destination = new MemoryStream(new byte[64], writable: true);

            OfficeStreamWriter.WriteAllBytes(destination, new byte[] { 1, 2, 3, 4 });

            Assert.Equal(0, destination.Position);
            Assert.Equal(4, destination.Length);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, destination.ToArray());
        }

        [Fact]
        public async Task WriteAllBytesAsyncTruncatesAndRewindsSeekableDestination() {
            using var destination = new MemoryStream(new byte[64], writable: true);

            await OfficeStreamWriter.WriteAllBytesAsync(
                destination,
                new byte[] { 5, 6, 7 },
                CancellationToken.None);

            Assert.Equal(0, destination.Position);
            Assert.Equal(3, destination.Length);
            Assert.Equal(new byte[] { 5, 6, 7 }, destination.ToArray());
        }

        [Fact]
        public void WriteAllBytesRejectsReadOnlyDestinationWithoutChangingIt() {
            using var destination = new MemoryStream(new byte[] { 9, 8, 7 }, writable: false);

            Assert.Throws<ArgumentException>(() =>
                OfficeStreamWriter.WriteAllBytes(destination, new byte[] { 1 }));

            Assert.Equal(3, destination.Length);
        }
    }
}
