using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class OfficeStreamReaderTests {
        [Fact]
        public void ReadRemainingBytesTreatsPositionPastEndAsEmpty() {
            using var source = new MemoryStream(new byte[] { 1, 2, 3 });
            source.Position = 8;

            byte[] result = OfficeStreamReader.ReadRemainingBytes(source, maxBytes: 1);

            Assert.Empty(result);
            Assert.Equal(8, source.Position);
        }
    }
}
