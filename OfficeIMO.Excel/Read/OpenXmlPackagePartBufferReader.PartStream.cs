#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class OpenXmlPackagePartBufferReader {
        /// <summary>
        /// Opens a part for forward-only reading with the same exact-length checks as a buffered read:
        /// a stream that ends early or continues past its declared length fails when the reader reaches it.
        /// </summary>
        internal bool TryOpenExactPartStream(string partName, CancellationToken ct, out Stream? stream) {
            stream = null;
            if (_disposed) throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            ct.ThrowIfCancellationRequested();
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out var entry) || entry.Length > int.MaxValue) return false;
            stream = new ExactLengthPartStream(OpenEntry(entry, ct), entry.Length, normalized, ct);
            return true;
        }

        private sealed class ExactLengthPartStream : Stream {
            private readonly Stream _input;
            private readonly long _length;
            private readonly string _partName;
            private readonly CancellationToken _ct;
            private long _read;
            private bool _endChecked;

            internal ExactLengthPartStream(Stream input, long length, string partName, CancellationToken ct) {
                _input = input;
                _length = length;
                _partName = partName;
                _ct = ct;
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => _length;
            public override long Position {
                get => _read;
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count) {
                _ct.ThrowIfCancellationRequested();
                long remaining = _length - _read;
                if (remaining == 0) {
                    if (!_endChecked) {
                        _endChecked = true;
                        if (_input.ReadByte() >= 0) {
                            throw ExcelPackagePartLengthFailure.Create(
                                $"Package part '{FormatPartNameForDisplay(_partName)}' exceeds its declared decompressed length of {_length} bytes.");
                        }
                    }
                    return 0;
                }
                if (count == 0) return 0;
                int read = _input.Read(buffer, offset, (int)Math.Min(count, remaining));
                if (read == 0) {
                    throw ExcelPackagePartLengthFailure.Create(
                        $"Package part '{FormatPartNameForDisplay(_partName)}' ended after {_read} of {_length} declared bytes.");
                }
                _read += read;
                return read;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (disposing) _input.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
