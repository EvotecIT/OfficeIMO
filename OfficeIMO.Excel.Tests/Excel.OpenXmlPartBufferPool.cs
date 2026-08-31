using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void OpenXmlPartBufferPool_RetainsLargePartsWithoutPowerOfTwoAmplification() {
            const int requestedLength = 34_402_657;
            byte[] first = OpenXmlPartBufferPool.Rent(requestedLength);
            try {
                Assert.InRange(first.Length, requestedLength, requestedLength + (64 * 1024) - 1);
                Assert.True(first.Length < 64 * 1024 * 1024);
                first[0] = 0x5A;
                first[requestedLength - 1] = 0xA5;
            } finally {
                OpenXmlPartBufferPool.Return(first);
            }

            byte[] reused = OpenXmlPartBufferPool.Rent(requestedLength);
            try {
                Assert.Same(first, reused);
                Assert.Equal(0, reused[0]);
                Assert.Equal(0, reused[requestedLength - 1]);
            } finally {
                OpenXmlPartBufferPool.Return(reused);
            }
        }

        [Fact]
        public void OpenXmlPartBufferPool_RejectsPartsAboveTheBoundedFastPath() {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                OpenXmlPartBufferPool.Rent((64 * 1024 * 1024) + 1));
        }
    }
}
