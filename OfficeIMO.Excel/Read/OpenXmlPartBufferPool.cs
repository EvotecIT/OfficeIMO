#nullable enable

namespace OfficeIMO.Excel {
    /// <summary>
    /// Retains a bounded set of large Open XML part buffers without the power-of-two
    /// amplification imposed by <see cref="System.Buffers.ArrayPool{T}.Shared"/>.
    /// </summary>
    internal static class OpenXmlPartBufferPool {
        private const int CapacityQuantum = 64 * 1024;
        private const int MaximumBufferSize = 64 * 1024 * 1024;
        private const long MaximumRetainedBytes = 64L * 1024 * 1024;
        private static readonly object Sync = new object();
        private static readonly Dictionary<int, byte[]> RetainedBuffers = new Dictionary<int, byte[]>();
        private static long _retainedBytes;

        internal static byte[] Rent(int minimumLength) {
            int capacity = GetCapacity(minimumLength);
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
            if (minimumLength < 0 || minimumLength > MaximumBufferSize) {
                throw new ArgumentOutOfRangeException(nameof(minimumLength));
            }

            int boundedLength = Math.Max(1, minimumLength);
            long capacity = ((long)boundedLength + CapacityQuantum - 1L) / CapacityQuantum * CapacityQuantum;
            return checked((int)capacity);
        }
    }
}
