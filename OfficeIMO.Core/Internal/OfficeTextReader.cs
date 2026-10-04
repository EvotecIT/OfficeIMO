using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>Decodes bounded text artifacts while preserving caller-owned stream state.</summary>
    internal static class OfficeTextReader {
        private const int BufferSize = 4096;

        /// <summary>Reads from the beginning, detects Unicode BOMs when no encoding is supplied, and limits decoded UTF-16 characters.</summary>
        public static string ReadAllText(Stream source, int? maximumCharacters, Encoding? encoding = null, CancellationToken cancellationToken = default) {
            Validate(source, maximumCharacters);
            cancellationToken.ThrowIfCancellationRequested();
            long position = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                using var reader = new StreamReader(new CancellationReadStream(source, cancellationToken), encoding ?? new UTF8Encoding(false), encoding == null, BufferSize, false);
                var output = new StringBuilder();
                var buffer = new char[BufferSize];
                while (true) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = reader.Read(buffer, 0, NextReadLength(output.Length, maximumCharacters));
                    cancellationToken.ThrowIfCancellationRequested();
                    if (read == 0) return output.ToString();
                    Append(output, buffer, read, maximumCharacters);
                }
            } finally {
                if (source.CanSeek) source.Seek(position, SeekOrigin.Begin);
            }
        }

        /// <summary>Asynchronously decodes bounded text without closing or repositioning the caller's stream permanently.</summary>
        public static async Task<string> ReadAllTextAsync(Stream source, int? maximumCharacters, Encoding? encoding = null, CancellationToken cancellationToken = default) {
            Validate(source, maximumCharacters);
            cancellationToken.ThrowIfCancellationRequested();
            long position = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                using var reader = new StreamReader(new CancellationReadStream(source, cancellationToken), encoding ?? new UTF8Encoding(false), encoding == null, BufferSize, false);
                var output = new StringBuilder();
                var buffer = new char[BufferSize];
                while (true) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int count = NextReadLength(output.Length, maximumCharacters);
#if NET6_0_OR_GREATER
                    int read = await reader.ReadAsync(buffer.AsMemory(0, count), cancellationToken).ConfigureAwait(false);
#else
                    int read = await reader.ReadAsync(buffer, 0, count).ConfigureAwait(false);
#endif
                    cancellationToken.ThrowIfCancellationRequested();
                    if (read == 0) return output.ToString();
                    Append(output, buffer, read, maximumCharacters);
                }
            } finally {
                if (source.CanSeek) source.Seek(position, SeekOrigin.Begin);
            }
        }

        private static int NextReadLength(int length, int? maximum) =>
            maximum.HasValue ? (int)Math.Min(BufferSize, (long)maximum.Value - length + 1) : BufferSize;

        private static void Append(StringBuilder output, char[] buffer, int count, int? maximum) {
            if (maximum.HasValue && (long)output.Length + count > maximum.Value) {
                throw new InvalidDataException("Text input exceeds the configured " + maximum.Value + " character limit.");
            }
            output.Append(buffer, 0, count);
        }

        private static void Validate(Stream source, int? maximum) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(source));
            if (maximum.HasValue && maximum.Value < 0) throw new ArgumentOutOfRangeException(nameof(maximum));
        }

        // Legacy StreamReader overloads do not accept a token. Forward the operation
        // token at the actual I/O boundary without disposing the caller-owned stream.
        private sealed class CancellationReadStream : Stream {
            private readonly Stream _source;
            private readonly CancellationToken _token;
            internal CancellationReadStream(Stream source, CancellationToken token) { _source = source; _token = token; }
            public override bool CanRead => _source.CanRead;
            public override bool CanSeek => _source.CanSeek;
            public override bool CanWrite => false;
            public override long Length => _source.Length;
            public override long Position { get => _source.Position; set => _source.Position = value; }
            public override int Read(byte[] buffer, int offset, int count) {
                _token.ThrowIfCancellationRequested();
                int result = _source.Read(buffer, offset, count);
                _token.ThrowIfCancellationRequested();
                return result;
            }
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
                _token.ThrowIfCancellationRequested();
                int result = await _source.ReadAsync(buffer, offset, count, _token).ConfigureAwait(false);
                _token.ThrowIfCancellationRequested();
                return result;
            }
            public override long Seek(long offset, SeekOrigin origin) => _source.Seek(offset, origin);
            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
