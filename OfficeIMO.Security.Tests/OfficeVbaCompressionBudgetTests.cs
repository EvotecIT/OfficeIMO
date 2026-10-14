using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security.Tests {
    public sealed class OfficeVbaCompressionBudgetTests {
        [Fact]
        public void FailedContainerChargesDiscardedExpansionWithoutReturningPartialSource() {
            byte[] compressed = OfficeVbaCompression.Compress(Enumerable.Repeat((byte)'a', 4096).ToArray()).Concat(new byte[] { 0 }).ToArray();
            int remaining = 5000;
            Assert.False(OfficeVbaCompression.TryDecompress(compressed, ref remaining, out byte[] source, out string detail));
            Assert.Empty(source); Assert.Contains("header", detail); Assert.Equal(904, remaining);
        }

        [Fact]
        public void CancelledExpansionPreservesTheUnspentBudget() {
            int remaining = 5000;
            Assert.Throws<OperationCanceledException>(() => OfficeVbaCompression.TryDecompress(new byte[] { 1 }, ref remaining, out _, out _, new CancellationToken(true)));
            Assert.Equal(5000, remaining);
        }
    }
}
