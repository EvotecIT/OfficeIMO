using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class OfficeStreamWriterTests {
        [Theory]
        [InlineData(false, new byte[] { 1, 2, 3, 4 })]
        [InlineData(true, new byte[] { 5, 6, 7 })]
        public async Task WriteAllBytesSyncAndAsyncTruncateAndRewindSeekableDestination(bool useAsync, byte[] bytes) {
            using var destination = new MemoryStream(new byte[64], writable: true);

            if (useAsync) {
                await OfficeStreamWriter.WriteAllBytesAsync(destination, bytes, CancellationToken.None);
            } else {
                OfficeStreamWriter.WriteAllBytes(destination, bytes);
            }

            Assert.Equal(0, destination.Position);
            Assert.Equal(bytes.Length, destination.Length);
            Assert.Equal(bytes, destination.ToArray());
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
