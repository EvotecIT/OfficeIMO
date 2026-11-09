#if NET8_0_OR_GREATER

namespace OfficeIMO.Excel {
    internal sealed partial class OpenXmlPackagePartBufferReader {
        /// <summary>Receives bulk BCL inflation into the existing bounded part buffer.</summary>
        private sealed class ExactPartCopyTarget : Stream {
            private readonly byte[] _output;
            private readonly int _length;
            private readonly string _partName;
            private readonly CancellationToken _cancellation;
            private readonly Action<int>? _publish;

            internal ExactPartCopyTarget(byte[] output, int offset, int length, string partName,
                CancellationToken cancellation, Action<int>? publish) {
                _output = output; Written = offset; _length = length; _partName = partName;
                _cancellation = cancellation; _publish = publish;
            }
            internal int Written { get; private set; }
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => Write(buffer.AsSpan(offset, count));
            public override void Write(ReadOnlySpan<byte> buffer) {
                _cancellation.ThrowIfCancellationRequested();
                if (buffer.Length > _length - Written) {
                    throw ExcelPackagePartLengthFailure.Create(
                        $"Package part '{FormatPartNameForDisplay(_partName)}' exceeds its declared decompressed length of {_length} bytes.");
                }
                buffer.CopyTo(_output.AsSpan(Written));
                Written += buffer.Length;
                _publish?.Invoke(Written);
            }
        }
    }
}
#endif
