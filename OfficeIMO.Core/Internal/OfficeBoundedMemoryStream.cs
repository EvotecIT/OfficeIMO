using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>Creates and identifies output-limit failures without conflating other invalid-data errors.</summary>
    internal static class OfficeOutputLimit {
        private const string Marker = "OfficeIMO.OutputLimit";

        internal static InvalidDataException Create(string message) {
            var exception = new InvalidDataException(message);
            exception.Data[Marker] = true;
            return exception;
        }

        internal static bool Is(Exception exception) => exception is InvalidDataException && exception.Data.Contains(Marker);
    }

    /// <summary>Bounds artifact serialization before writes or explicit capacity changes grow the buffer.</summary>
    internal class OfficeBoundedMemoryStream : MemoryStream {
        private readonly long _maximumBytes;

        internal OfficeBoundedMemoryStream(long maximumBytes, int capacityHint = 0)
            : base(GetInitialCapacity(maximumBytes, capacityHint)) {
            _maximumBytes = maximumBytes;
        }

        public override int Capacity {
            get => base.Capacity;
            set { EnsureLength(value); base.Capacity = value; }
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (offset > buffer.Length - count) throw new ArgumentException("Offset and count exceed the buffer.");
            EnsureWrite(count);
            base.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) { EnsureWrite(1); base.WriteByte(value); }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            Write(buffer, offset, count);
            return Task.CompletedTask;
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) { EnsureWrite(buffer.Length); base.Write(buffer); }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            Write(buffer.Span);
            return default;
        }
#endif

        public override void SetLength(long value) {
            EnsureLength(value);
            EnsureCapacityWithinLimit(value);
            base.SetLength(value);
        }

        protected virtual Exception CreateLimitException(long maximumBytes) =>
            OfficeOutputLimit.Create($"The artifact exceeds the configured output limit of {maximumBytes} bytes.");

        private void EnsureWrite(int count) {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            long end;
            try { end = checked(Position + count); }
            catch (OverflowException) { throw CreateLimitException(_maximumBytes); }
            EnsureLength(Math.Max(Length, end));
            EnsureCapacityWithinLimit(end);
        }

        private void EnsureLength(long value) {
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            if (value > _maximumBytes) throw CreateLimitException(_maximumBytes);
        }

        private void EnsureCapacityWithinLimit(long end) {
            if (end <= base.Capacity) return;
            base.Capacity = checked((int)Math.Min(_maximumBytes, Math.Max(end, Math.Max(256L, base.Capacity * 2L))));
        }

        private static int GetInitialCapacity(long maximumBytes, int capacityHint) {
            if (maximumBytes <= 0 || maximumBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            if (capacityHint < 0) throw new ArgumentOutOfRangeException(nameof(capacityHint));
            return (int)Math.Min(maximumBytes, capacityHint);
        }
    }
}
