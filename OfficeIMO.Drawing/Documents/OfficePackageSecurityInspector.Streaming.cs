using System;
using System.IO;

namespace OfficeIMO.Drawing {
    public static partial class OfficePackageSecurityInspector {
        private static OfficePackageSecurityReport InspectSeekableSource(
            Stream source,
            OfficePackageSecurityOptions options) {
            ValidateOptions(options);
            if (!source.CanRead || !source.CanSeek) {
                throw new ArgumentException(
                    "Streaming package inspection requires a readable seekable stream.",
                    nameof(source));
            }

            long originalPosition = source.Position;
            try {
                long packageBytes = checked(source.Length - originalPosition);
                ValidateSourceSize(packageBytes, options);
                var findings = new System.Collections.Generic.List<OfficePackageSecurityFinding>();

                using var package = new SeekableReadWindowStream(source, originalPosition, packageBytes);
                var signature = new byte[8];
                int signatureBytes = ReadPrefix(package, signature);
                package.Position = 0;
                bool isZip = HasZipSignature(signature, signatureBytes);
                bool isCompound = OfficeIMO.Drawing.Internal.OfficeCompoundDocumentDetector
                    .HasCompoundSignature(signature);
                if (isZip) return InspectZip(package, packageBytes, options, findings);
                if (isCompound) return InspectCompound(package, packageBytes, options, findings);
                return new OfficePackageSecurityReport(packageBytes, OfficePackageContainerKind.Unknown,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            } finally {
                source.Position = originalPosition;
            }
        }

        private static int ReadPrefix(Stream source, byte[] buffer) {
            int total = 0;
            while (total < buffer.Length) {
                int read = source.Read(buffer, total, buffer.Length - total);
                if (read == 0) break;
                total += read;
            }
            return total;
        }

        private static bool HasZipSignature(byte[] bytes, int length) => length >= 4
            && bytes[0] == 0x50 && bytes[1] == 0x4b
            && ((bytes[2] == 0x03 && bytes[3] == 0x04)
                || (bytes[2] == 0x05 && bytes[3] == 0x06)
                || (bytes[2] == 0x07 && bytes[3] == 0x08));

        private sealed class SeekableReadWindowStream : Stream {
            private readonly Stream _source;
            private readonly long _offset;
            private readonly long _length;
            private long _position;

            internal SeekableReadWindowStream(Stream source, long offset, long length) {
                _source = source;
                _offset = offset;
                _length = length;
            }

            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length => _length;
            public override long Position {
                get => _position;
                set => Seek(value, SeekOrigin.Begin);
            }

            public override int Read(byte[] buffer, int offset, int count) {
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                    throw new ArgumentOutOfRangeException(nameof(offset));
                }
                long remaining = _length - _position;
                if (remaining <= 0) return 0;
                int requested = (int)Math.Min(count, remaining);
                _source.Position = checked(_offset + _position);
                int read = _source.Read(buffer, offset, requested);
                _position += read;
                return read;
            }

            public override long Seek(long offset, SeekOrigin origin) {
                long target = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked(_position + offset),
                    SeekOrigin.End => checked(_length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                if (target < 0 || target > _length) throw new IOException(
                    "Attempted to seek outside the package inspection window.");
                _position = target;
                return target;
            }

            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
