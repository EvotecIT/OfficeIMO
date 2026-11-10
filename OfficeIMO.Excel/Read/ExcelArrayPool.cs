using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Worksheet index and part-writing arrays. Small arrays use the shared pool; an array sized for a whole large
    /// worksheet is kept only until the collector next needs the memory, so it does not stay in the process after the
    /// read or save has finished.
    /// </summary>
    internal static class ExcelArrayPool {
        private const int MaximumSharedBytes = 1024 * 1024;
        private const int MaximumPerLength = 8;
        private const long MaximumRetainedBytes = 256L * 1024 * 1024;
        private static readonly ExcelReleasingArrayStore<int> LargeInts =
            new ExcelReleasingArrayStore<int>(MaximumPerLength, MaximumRetainedBytes / sizeof(int));
        private static readonly ExcelReleasingArrayStore<byte> LargeBytes =
            new ExcelReleasingArrayStore<byte>(MaximumPerLength, MaximumRetainedBytes);
        private static readonly ExcelReleasingArrayStore<long> LargeLongs =
            new ExcelReleasingArrayStore<long>(MaximumPerLength, MaximumRetainedBytes / sizeof(long));

        internal static int[] RentInts(int minimumLength) => RentInts(minimumLength, releaseWhenIdle: false);

        /// <summary>
        /// Rents ints. <paramref name="releaseWhenIdle"/> keeps even a small array out of the shared pool, for storage
        /// that only lives during one read, such as the many chunks of a parallel worksheet index.
        /// </summary>
        internal static int[] RentInts(int minimumLength, bool releaseWhenIdle) {
            if (!releaseWhenIdle && minimumLength <= MaximumSharedBytes / sizeof(int)) return ArrayPool<int>.Shared.Rent(minimumLength);
            int length = RoundUpToPowerOfTwo(minimumLength);
            return LargeInts.TryTake(length) ?? new int[length];
        }

        internal static byte[] RentBytes(int minimumLength) => RentBytes(minimumLength, releaseWhenIdle: false);

        internal static byte[] RentBytes(int minimumLength, bool releaseWhenIdle) {
            if (!releaseWhenIdle && minimumLength <= MaximumSharedBytes) return ArrayPool<byte>.Shared.Rent(minimumLength);
            int length = RoundUpToPowerOfTwo(minimumLength);
            return LargeBytes.TryTake(length) ?? new byte[length];
        }

        internal static long[] RentLongs(int minimumLength) => RentLongs(minimumLength, releaseWhenIdle: false);

        internal static long[] RentLongs(int minimumLength, bool releaseWhenIdle) {
            if (!releaseWhenIdle && minimumLength <= MaximumSharedBytes / sizeof(long)) return ArrayPool<long>.Shared.Rent(minimumLength);
            int length = RoundUpToPowerOfTwo(minimumLength);
            return LargeLongs.TryTake(length) ?? new long[length];
        }

        internal static void Return(int[] array) {
            if (array.Length <= MaximumSharedBytes / sizeof(int)) ArrayPool<int>.Shared.Return(array);
            else LargeInts.TryAdd(array);
        }

        /// <summary>Returns an array rented with <c>releaseWhenIdle: true</c>.</summary>
        internal static void ReturnIdleReleased(int[] array) => LargeInts.TryAdd(array);

        internal static void ReturnIdleReleased(byte[] array) => LargeBytes.TryAdd(array);

        internal static void ReturnIdleReleased(long[] array) => LargeLongs.TryAdd(array);

        internal static void Return(byte[] array) => Return(array, clearArray: false);

        /// <summary>Returns bytes, clearing them first when they held workbook content.</summary>
        internal static void Return(byte[] array, bool clearArray) {
            if (array.Length <= MaximumSharedBytes) {
                ArrayPool<byte>.Shared.Return(array, clearArray);
                return;
            }
            if (clearArray) Array.Clear(array, 0, array.Length);
            LargeBytes.TryAdd(array);
        }

        internal static void Return(long[] array) {
            if (array.Length <= MaximumSharedBytes / sizeof(long)) ArrayPool<long>.Shared.Return(array);
            else LargeLongs.TryAdd(array);
        }

        private static int RoundUpToPowerOfTwo(int value) {
            if (value > 1 << 30) return value;
            int result = 1;
            while (result < value) result <<= 1;
            return result;
        }
    }
}
