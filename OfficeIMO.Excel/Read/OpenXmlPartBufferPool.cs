namespace OfficeIMO.Excel {
    /// <summary>
    /// Retains a bounded set of large Open XML part buffers without the power-of-two
    /// amplification imposed by <see cref="System.Buffers.ArrayPool{T}.Shared"/>.
    /// </summary>
    internal static class OpenXmlPartBufferPool {
        private const int CapacityQuantum = 64 * 1024;
        private const int MaximumBufferSize = 256 * 1024 * 1024;
        private const long MaximumRetainedBytes = 64L * 1024 * 1024;
        // A buffer for a whole large part helps the next read in a batch, but keeping it after the work is done would
        // hold that much memory for as long as the process runs, so buffers above this size are only kept until the
        // collector needs the memory.
        private const int MaximumLongLivedBufferSize = 4 * 1024 * 1024;
        private static readonly object Sync = new object();
        private static readonly Dictionary<int, byte[]> RetainedBuffers = new Dictionary<int, byte[]>();
        private static readonly ExcelReleasingArrayStore<byte> LargeBuffers = new ExcelReleasingArrayStore<byte>(2, MaximumRetainedBytes);
        private static long _retainedBytes;

        internal static byte[] Rent(int minimumLength) {
            int capacity = GetCapacity(minimumLength);
            if (capacity > MaximumLongLivedBufferSize) {
                return LargeBuffers.TryTake(capacity) ?? new byte[capacity];
            }

            lock (Sync) {
                if (RetainedBuffers.TryGetValue(capacity, out byte[]? retained)) {
                    RetainedBuffers.Remove(capacity);
                    _retainedBytes -= capacity;
                    return retained;
                }
            }

            return new byte[capacity];
        }

        internal static void Return(byte[]? buffer, bool retain = true) {
            if (buffer == null) return;

            // Workbook XML may contain sensitive customer data. Clear the complete
            // capacity before either retaining the buffer or releasing it to the GC.
            Array.Clear(buffer, 0, buffer.Length);
            if (!retain
                || buffer.Length < CapacityQuantum
                || buffer.Length > MaximumBufferSize
                || buffer.Length % CapacityQuantum != 0) {
                return;
            }

            if (buffer.Length > MaximumLongLivedBufferSize) {
                LargeBuffers.TryAdd(buffer);
                return;
            }

            lock (Sync) {
                if (RetainedBuffers.ContainsKey(buffer.Length)
                    || _retainedBytes + buffer.Length > MaximumRetainedBytes) {
                    return;
                }

                RetainedBuffers.Add(buffer.Length, buffer);
                _retainedBytes += buffer.Length;
            }
        }

        internal static int GetCapacity(int minimumLength) {
            ArgumentOutOfRangeException.ThrowIfNegative(minimumLength);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(minimumLength, MaximumBufferSize);

            int boundedLength = Math.Max(1, minimumLength);
            long capacity = ((long)boundedLength + CapacityQuantum - 1L) / CapacityQuantum * CapacityQuantum;
            return checked((int)capacity);
        }
    }
}
