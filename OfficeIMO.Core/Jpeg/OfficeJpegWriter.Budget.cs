using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegWriter {
    private static long GetCoefficientStorageBytes(
        int width,
        int height,
        ComponentSpec[] components,
        int maxH,
        int maxV) {
        try {
            long mcuCols = ((long)width + maxH * 8L - 1L) / (maxH * 8L);
            long mcuRows = ((long)height + maxV * 8L - 1L) / (maxV * 8L);
            long bytes = 24L + components.LongLength * 32L;
            for (int index = 0; index < components.Length; index++) {
                ComponentSpec component = components[index];
                long elements = checked(
                    mcuCols * component.H * mcuRows * component.V * 64L);
                if (elements > int.MaxValue) throw new ArgumentException(JpegOutputLimitMessage);
                bytes = checked(bytes + elements * sizeof(short) + 24L);
            }
            return bytes;
        } catch (OverflowException) {
            throw new ArgumentException(JpegOutputLimitMessage);
        }
    }

    private static long GetMetadataManagedBytes(OfficeJpegMetadata metadata) {
        long bytes = 0L;
        if (metadata.ExifBuffer != null) bytes = checked(bytes + metadata.ExifBuffer.LongLength + 24L);
        if (metadata.XmpBuffer != null) bytes = checked(bytes + metadata.XmpBuffer.LongLength + 24L);
        if (metadata.IccBuffer != null) bytes = checked(bytes + metadata.IccBuffer.LongLength + 24L);
        return bytes;
    }

    private static long GetFixedEncodingManagedBytes(
        long rgbaBytes,
        long coefficientBytes,
        long metadataBytes,
        long retainedManagedBytes) {
        if (rgbaBytes < 0L || coefficientBytes < 0L || metadataBytes < 0L || retainedManagedBytes < 0L) {
            throw new ArgumentException(JpegOutputLimitMessage);
        }
        try {
            return checked(
                rgbaBytes + 24L + coefficientBytes + metadataBytes +
                retainedManagedBytes + JpegWorkingScratchBytes);
        } catch (OverflowException) {
            throw new ArgumentException(JpegOutputLimitMessage);
        }
    }

    internal static bool IsEncodingWorkingSetWithinLimit(
        long fixedManagedBytes,
        long outputBackingBytes,
        long additionalOutputBytes = 0L) {
        if (fixedManagedBytes < 0L || outputBackingBytes < 0L || additionalOutputBytes < 0L ||
            outputBackingBytes > OfficeRasterGuards.MaximumEncodedBytes) return false;
        try {
            return checked(
                       fixedManagedBytes + outputBackingBytes + 24L + additionalOutputBytes) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool IsEncodingGrowthWorkingSetWithinLimit(
        long fixedManagedBytes,
        long currentOutputBackingBytes,
        long newOutputBackingBytes) {
        if (fixedManagedBytes < 0L || currentOutputBackingBytes < 0L || newOutputBackingBytes < 0L ||
            newOutputBackingBytes > OfficeRasterGuards.MaximumEncodedBytes) return false;
        try {
            return checked(
                       fixedManagedBytes + currentOutputBackingBytes + 24L +
                       newOutputBackingBytes + 24L) <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static void EnsureEncodingWorkingSet(long fixedManagedBytes, long outputBackingBytes) {
        if (!IsEncodingWorkingSetWithinLimit(fixedManagedBytes, outputBackingBytes)) {
            throw new ArgumentException(JpegOutputLimitMessage);
        }
    }

    private static long GetMemoryStreamBackingBytes(Stream stream) {
        if (!(stream is MemoryStream memoryStream)) return 0L;
        if (memoryStream.TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null) {
            return segment.Array.LongLength;
        }
        if (memoryStream.Capacity == 0) return 0L;
        throw new ArgumentException(
            "JPEG output MemoryStream must expose its retained buffer for bounded encoding.", nameof(stream));
    }

    private sealed class JpegBudgetedMemoryStream : Stream {
        private readonly MemoryStream _inner;
        private readonly long _fixedManagedBytes;

        internal JpegBudgetedMemoryStream(MemoryStream inner, long fixedManagedBytes) {
            _inner = inner;
            _fixedManagedBytes = fixedManagedBytes;
        }

        public override bool CanRead => false;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        public override void SetLength(long value) {
            EnsureCapacityFor(value);
            _inner.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }
            EnsureCapacityFor(checked(_inner.Position + count));
            _inner.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) {
            EnsureCapacityFor(checked(_inner.Position + 1L));
            _inner.WriteByte(value);
        }

        private void EnsureCapacityFor(long requiredPosition) {
            long requiredLength = Math.Max(requiredPosition, _inner.Length);
            if (requiredLength > OfficeRasterGuards.MaximumEncodedBytes) {
                throw new ArgumentException(JpegOutputLimitMessage);
            }
            long currentBackingBytes = GetMemoryStreamBackingBytes(_inner);
            if (requiredLength > _inner.Capacity) {
                long doubled = Math.Max(256L, checked((long)_inner.Capacity * 2L));
                long projectedBackingBytes = Math.Max(requiredLength, doubled);
                if (!IsEncodingGrowthWorkingSetWithinLimit(
                        _fixedManagedBytes, currentBackingBytes, projectedBackingBytes)) {
                    throw new ArgumentException(JpegOutputLimitMessage);
                }
                return;
            }
            EnsureEncodingWorkingSet(_fixedManagedBytes, currentBackingBytes);
        }
    }
}
