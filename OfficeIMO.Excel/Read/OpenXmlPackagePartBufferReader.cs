#nullable enable

using System.Buffers;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Provides bounded, exact-size reads for package parts whose OPC streams do not expose
    /// their declared uncompressed length. This avoids repeated growth copies on large sheets.
    /// </summary>
    internal sealed class OpenXmlPackagePartBufferReader : IDisposable {
        private readonly Stream _stream;
        private readonly ZipArchive _archive;
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private readonly object _prefetchSync = new object();
        private Task<PrefetchedPartBuffer>? _prefetchTask;
        private CancellationTokenSource? _prefetchCancellation;
        private string? _prefetchPartName;
        private bool _disposed;

        private OpenXmlPackagePartBufferReader(Stream stream) {
            _stream = stream;
            _archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try {
                _entries = BuildEntryIndex(_archive);
            } catch {
                _archive.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Opens a ZIP reader that takes ownership of the supplied stream, including when
        /// the stream does not contain a readable ZIP archive.
        /// </summary>
        internal static OpenXmlPackagePartBufferReader? TryOpen(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            try {
                var reader = new OpenXmlPackagePartBufferReader(stream);
                return reader;
            } catch (OpenXmlPackagePartIndexException) {
                stream.Dispose();
                throw;
            } catch (InvalidDataException) {
                stream.Dispose();
                return null;
            } catch (IOException) {
                stream.Dispose();
                return null;
            } catch (UnauthorizedAccessException) {
                stream.Dispose();
                return null;
            } catch (NotSupportedException) {
                stream.Dispose();
                return null;
            } catch (ArgumentException) {
                stream.Dispose();
                return null;
            }
        }

        internal static OpenXmlPackagePartBufferReader? TryOpen(byte[] bytes) {
            MemoryStream? stream = null;
            try {
                stream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false);
                var reader = new OpenXmlPackagePartBufferReader(stream);
                stream = null;
                return reader;
            } catch (OpenXmlPackagePartIndexException) {
                stream?.Dispose();
                throw;
            } catch (InvalidDataException) {
                stream?.Dispose();
                return null;
            } catch (ArgumentException) {
                stream?.Dispose();
                return null;
            }
        }

        internal bool TryRead(
            Uri partUri,
            int maximumBytes,
            CancellationToken cancellationToken,
            out byte[]? buffer,
            out int length) {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            }
            cancellationToken.ThrowIfCancellationRequested();
            buffer = null;
            length = 0;
            return TryRead(
                partUri.OriginalString,
                maximumBytes,
                cancellationToken,
                out buffer,
                out length);
        }

        internal bool TryRead(
            string partName,
            int maximumBytes,
            CancellationToken cancellationToken,
            out byte[]? buffer,
            out int length) {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            }
            if (maximumBytes < 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }
            cancellationToken.ThrowIfCancellationRequested();
            buffer = null;
            length = 0;
            string normalizedPartName = NormalizePartName(partName);
            _entries.TryGetValue(normalizedPartName, out ZipArchiveEntry? entry);
            if (entry == null || entry.Length < 0 || entry.Length > maximumBytes || entry.Length > int.MaxValue) {
                return false;
            }

            Task<PrefetchedPartBuffer>? prefetchTask = null;
            CancellationTokenSource? prefetchCancellation = null;
            lock (_prefetchSync) {
                if (string.Equals(_prefetchPartName, normalizedPartName, StringComparison.OrdinalIgnoreCase)) {
                    prefetchTask = _prefetchTask;
                    prefetchCancellation = _prefetchCancellation;
                    _prefetchTask = null;
                    _prefetchCancellation = null;
                    _prefetchPartName = null;
                }
            }

            PrefetchedPartBuffer? prefetched;
            if (prefetchTask != null) {
                try {
                    prefetched = prefetchTask.GetAwaiter().GetResult();
                } finally {
                    prefetchCancellation?.Dispose();
                }
            } else {
                prefetched = ReadPart(entry, normalizedPartName, cancellationToken);
            }
            if (prefetched == null) return false;
            prefetched.Detach(out buffer, out length);
            return true;
        }

        internal void BeginPrefetch(
            string partName,
            int maximumBytes,
            CancellationToken cancellationToken) {
            string normalizedPartName = NormalizePartName(partName);
            if (_disposed) throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            cancellationToken.ThrowIfCancellationRequested();
            if (!_entries.TryGetValue(normalizedPartName, out ZipArchiveEntry? entry)
                || entry.Length < 0
                || entry.Length > maximumBytes
                || entry.Length > int.MaxValue) {
                return;
            }

            lock (_prefetchSync) {
                if (_prefetchTask != null) {
                    throw new InvalidOperationException("A package-part prefetch is already active.");
                }
                var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                _prefetchCancellation = linked;
                _prefetchPartName = normalizedPartName;
                _prefetchTask = Task.Run(
                    () => ReadPart(entry, normalizedPartName, linked.Token),
                    CancellationToken.None);
            }
        }

        private static PrefetchedPartBuffer ReadPart(
            ZipArchiveEntry entry,
            string normalizedPartName,
            CancellationToken cancellationToken) {
            int length = checked((int)entry.Length);
            byte[] output = ArrayPool<byte>.Shared.Rent(Math.Max(1, length));
            try {
                using Stream input = entry.Open();
                int offset = 0;
                while (offset < length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = input.Read(output, offset, length - offset);
                    if (read == 0) {
                        throw new EndOfStreamException(
                            $"Package part '{FormatPartNameForDisplay(normalizedPartName)}' ended after {offset} of {length} declared bytes.");
                    }
                    offset += read;
                }
                cancellationToken.ThrowIfCancellationRequested();
                if (input.ReadByte() >= 0) {
                    throw new InvalidDataException(
                        $"Package part '{FormatPartNameForDisplay(normalizedPartName)}' exceeds its declared decompressed length of {length} bytes.");
                }
                return new PrefetchedPartBuffer(output, length);
            } catch {
                ArrayPool<byte>.Shared.Return(output);
                throw;
            }
        }

        internal bool ContainsPart(string partName) {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            }

            return _entries.ContainsKey(NormalizePartName(partName));
        }

        internal Stream OpenPart(string partName, int maximumBytes) {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
            }
            if (maximumBytes < 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }

            string normalizedPartName = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalizedPartName, out ZipArchiveEntry? entry)) {
                throw new InvalidDataException($"Package part '{FormatPartNameForDisplay(normalizedPartName)}' is missing.");
            }
            if (entry.Length < 0 || entry.Length > maximumBytes || entry.Length > int.MaxValue) {
                throw ExcelReadLimitFailure.Create(
                    $"Package part '{FormatPartNameForDisplay(normalizedPartName)}' declares {entry.Length} bytes, exceeding the supported limit of {maximumBytes} bytes.");
            }

            return entry.Open();
        }

        private static Dictionary<string, ZipArchiveEntry> BuildEntryIndex(ZipArchive archive) {
            var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries) {
                if (string.IsNullOrEmpty(entry.Name)) {
                    continue;
                }

                string normalized = NormalizePartName(entry.FullName);
                if (entries.ContainsKey(normalized)) {
                    throw new OpenXmlPackagePartIndexException(
                        $"The Open XML package contains ambiguous part name '{FormatPartNameForDisplay(normalized)}'.");
                }

                entries.Add(normalized, entry);
            }

            return entries;
        }

        private static string NormalizePartName(string partName) {
            if (string.IsNullOrWhiteSpace(partName)) {
                throw new ArgumentException("Package part name cannot be empty.", nameof(partName));
            }
            if (partName.IndexOf('\\') >= 0) {
                throw new OpenXmlPackagePartIndexException($"Package part name '{partName}' contains a backslash.");
            }

            string normalized = partName.TrimStart('/');
            string[] encodedSegments = normalized.Split('/');
            var canonicalSegments = new string[encodedSegments.Length];
            for (int index = 0; index < encodedSegments.Length; index++) {
                string encodedSegment = encodedSegments[index];
                if (encodedSegment.Length == 0) {
                    throw new OpenXmlPackagePartIndexException($"Package part name '{partName}' contains an empty segment.");
                }

                for (int characterIndex = 0; characterIndex < encodedSegment.Length; characterIndex++) {
                    if (encodedSegment[characterIndex] != '%') continue;
                    if (characterIndex + 2 >= encodedSegment.Length
                        || !IsHexDigit(encodedSegment[characterIndex + 1])
                        || !IsHexDigit(encodedSegment[characterIndex + 2])) {
                        throw new OpenXmlPackagePartIndexException($"Package part name '{partName}' contains invalid percent encoding.");
                    }
                    characterIndex += 2;
                }

                string decodedSegment;
                try {
                    decodedSegment = Uri.UnescapeDataString(encodedSegment);
                } catch (UriFormatException exception) {
                    throw new OpenXmlPackagePartIndexException(
                        $"Package part name '{partName}' contains invalid percent encoding.",
                        exception);
                }
                if (decodedSegment.Length == 0
                    || decodedSegment == "."
                    || decodedSegment == ".."
                    || decodedSegment.IndexOf('/') >= 0
                    || decodedSegment.IndexOf('\\') >= 0
                    || decodedSegment.Any(char.IsControl)) {
                    throw new OpenXmlPackagePartIndexException($"Package part name '{partName}' is not safe.");
                }

                // OPC relationship targets are URI-encoded while ZIP writers differ on
                // whether FullName exposes the encoded or decoded spelling. Re-encode once
                // to an idempotent URI key so %20 and %2520 remain distinct.
                canonicalSegments[index] = Uri.EscapeDataString(decodedSegment);
            }

            return string.Join("/", canonicalSegments);
        }

        private static string FormatPartNameForDisplay(string canonicalPartName) {
            try {
                return Uri.UnescapeDataString(canonicalPartName);
            } catch (UriFormatException) {
                return canonicalPartName;
            }
        }

        private static bool IsHexDigit(char value) =>
            value is >= '0' and <= '9'
            || value is >= 'a' and <= 'f'
            || value is >= 'A' and <= 'F';

        public void Dispose() {
            if (_disposed) {
                return;
            }
            _disposed = true;
            Task<PrefetchedPartBuffer>? prefetchTask;
            CancellationTokenSource? prefetchCancellation;
            lock (_prefetchSync) {
                prefetchTask = _prefetchTask;
                prefetchCancellation = _prefetchCancellation;
                _prefetchTask = null;
                _prefetchCancellation = null;
                _prefetchPartName = null;
            }
            prefetchCancellation?.Cancel();
            try {
                prefetchTask?.GetAwaiter().GetResult()?.Dispose();
            } catch {
                // Disposal observes unconsumed canceled or faulted prefetch work.
            } finally {
                prefetchCancellation?.Dispose();
            }
            _archive.Dispose();
            _stream.Dispose();
        }

        private sealed class PrefetchedPartBuffer : IDisposable {
            private byte[]? _buffer;

            internal PrefetchedPartBuffer(byte[] buffer, int length) {
                _buffer = buffer;
                Length = length;
            }

            internal int Length { get; }

            internal void Detach(out byte[] buffer, out int length) {
                buffer = _buffer ?? throw new ObjectDisposedException(nameof(PrefetchedPartBuffer));
                _buffer = null;
                length = Length;
            }

            public void Dispose() {
                byte[]? buffer = _buffer;
                _buffer = null;
                if (buffer != null) ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }

    internal sealed class OpenXmlPackagePartIndexException : IOException {
        internal OpenXmlPackagePartIndexException(string message) : base(message) { }

        internal OpenXmlPackagePartIndexException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
