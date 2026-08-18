#nullable enable

using System.Buffers;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Provides bounded, exact-size reads for package parts whose OPC streams do not expose
    /// their declared uncompressed length. This avoids repeated growth copies on large sheets.
    /// </summary>
    internal sealed class OpenXmlPackagePartBufferReader : IDisposable {
        private readonly Stream _stream;
        private readonly ZipArchive _archive;
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private bool _disposed;

        private OpenXmlPackagePartBufferReader(Stream stream) {
            _stream = stream;
            _archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            _entries = BuildEntryIndex(_archive);
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
            }
        }

        internal static OpenXmlPackagePartBufferReader? TryOpen(byte[] bytes) {
            MemoryStream? stream = null;
            try {
                stream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: false);
                var reader = new OpenXmlPackagePartBufferReader(stream);
                stream = null;
                return reader;
            } catch (InvalidDataException) {
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

            length = checked((int)entry.Length);
            byte[] output = ArrayPool<byte>.Shared.Rent(Math.Max(1, length));
            try {
                using Stream input = entry.Open();
                int offset = 0;
                while (offset < length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = input.Read(output, offset, length - offset);
                    if (read == 0) {
                        throw new EndOfStreamException(
                            $"Package part '{normalizedPartName}' ended after {offset} of {length} declared bytes.");
                    }
                    offset += read;
                }
                cancellationToken.ThrowIfCancellationRequested();
                if (input.ReadByte() >= 0) {
                    throw new InvalidDataException(
                        $"Package part '{normalizedPartName}' exceeds its declared decompressed length of {length} bytes.");
                }
                buffer = output;
                return true;
            } catch {
                ArrayPool<byte>.Shared.Return(output);
                buffer = null;
                length = 0;
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
                throw new InvalidDataException($"Package part '{normalizedPartName}' is missing.");
            }
            if (entry.Length < 0 || entry.Length > maximumBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException(
                    $"Package part '{normalizedPartName}' declares {entry.Length} bytes, exceeding the supported limit of {maximumBytes} bytes.");
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
                    throw new InvalidDataException($"The Open XML package contains duplicate part name '{normalized}'.");
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
                throw new InvalidDataException($"Package part name '{partName}' contains a backslash.");
            }

            string normalized = partName.TrimStart('/');
            if (normalized.Split('/').Any(static segment => segment == "..")) {
                throw new InvalidDataException($"Package part name '{partName}' is not safe.");
            }

            return normalized;
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }
            _disposed = true;
            _archive.Dispose();
            _stream.Dispose();
        }
    }
}
