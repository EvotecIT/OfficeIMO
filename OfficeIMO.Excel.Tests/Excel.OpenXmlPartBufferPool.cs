using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void OpenXmlPartBufferPool_BoundsCapacityAndClearsDiscardedStorage() {
            const int requestedLength = 34_402_657;
            byte[] first = OpenXmlPartBufferPool.Rent(requestedLength);
            try {
                Assert.InRange(first.Length, requestedLength, requestedLength + (64 * 1024) - 1);
                first[0] = 0x5A;
                first[requestedLength - 1] = 0xA5;
                first[first.Length - 1] = 0xFF;
            } finally {
                // Retention is optional and can decline when other reads fill the pool.
                // Discarding keeps this buffer private while its clearing is verified.
                OpenXmlPartBufferPool.Return(first, retain: false);
            }

            Assert.Equal(0, first[0]);
            Assert.Equal(0, first[requestedLength - 1]);
            Assert.Equal(0, first[first.Length - 1]);
        }

        [Fact]
        public void OpenXmlPartBufferPool_RejectsPartsAboveTheBoundedFastPath() {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                OpenXmlPartBufferPool.Rent((64 * 1024 * 1024) + 1));
        }
    }
}
