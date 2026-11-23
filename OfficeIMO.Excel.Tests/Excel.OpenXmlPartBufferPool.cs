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
        public void OpenXmlPartBufferPool_ReusesWholePartBuffersAboveTheLongLivedLimit() {
            // A worksheet of about 1M typed rows inflates to roughly this size. Reads in a batch must reuse its buffer
            // rather than allocate the whole part again on each read.
            // A fresh store with the pool's limits keeps the result independent of buffers other tests leave retained.
            const int requestedLength = 170_561_745;
            var store = OpenXmlPartBufferPool.CreateLargeBufferStore();
            byte[] buffer = new byte[OpenXmlPartBufferPool.GetCapacity(requestedLength)];
            Assert.True(store.TryAdd(buffer));
            Assert.Same(buffer, store.TryTake(buffer.Length));
        }

        [Fact]
        public void OpenXmlPartBufferPool_RejectsPartsAboveTheBoundedFastPath() {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                OpenXmlPartBufferPool.Rent((256 * 1024 * 1024) + 1));
        }
    }
}
