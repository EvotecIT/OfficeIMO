using System.IO.Compression;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    internal sealed partial class OpenXmlPackagePartBufferReader {
        private List<OfficeArchiveSafety.ZipCentralDirectoryRecord>? _zipCoordinates;
        private bool _checkedZipCoordinates;

        private OpenXmlPackagePartBufferReader(Stream stream, OpenXmlPackagePartBufferReader source, bool allowPipelining) {
            _stream = stream;
            _entries = source._entries;
            _archiveEntryCount = source._archiveEntryCount;
            _sharedEntryIndex = true;
            AllowsPipelining = allowPipelining;
        }

        /// <summary>Shares immutable entry metadata while keeping the same snapshot's worksheet cursor independent.</summary>
        internal OpenXmlPackagePartBufferReader? TryOpenIndependent(Stream stream, bool allowPipelining = false) {
            try {
                if (_disposed) throw new ObjectDisposedException(nameof(OpenXmlPackagePartBufferReader));
                // .NET 11 adds CRC checks to ZIP entry streams; retain that authoritative reader.
                if (Environment.Version.Major >= 11 || !_stream.CanSeek || !stream.CanSeek
                    || stream.Length != _stream.Length) return TryOpen(stream, allowPipelining);
                return new OpenXmlPackagePartBufferReader(stream, this, allowPipelining);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>Reads one independently bounded payload, retaining BCL entry reads for unsupported ZIP layouts.</summary>
        private Stream OpenEntry(ZipArchiveEntry entry, CancellationToken cancellationToken) {
            if (!_sharedEntryIndex) return entry.Open();
            if (TryGetZipCoordinates(entry, cancellationToken, out var record)) {
                try {
                    long payload = OfficeArchiveSafety.ValidateEntryPayload(_stream, record);
                    var window = new SeekableReadWindowStream(_stream, payload, record.CompressedSize);
                    if (record.CompressionMethod == 0) return window;
                    try {
                        // The caller consumes the entire part and checks both its exact expanded
                        // length and EOF, just as it does for the corresponding ZipArchive stream.
                        return new DeflateStream(window, CompressionMode.Decompress);
                    } catch {
                        window.Dispose();
                        throw;
                    }
                } catch (InvalidDataException) {
                    // Preserve the original archive's behavior for layouts outside the
                    // stricter shared ZIP-coordinate proof.
                }
            }

            _archive ??= new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: true);
            return (_archive.GetEntry(entry.FullName)
                ?? throw new InvalidDataException("An indexed package entry is missing from its independent snapshot.")).Open();
        }

        /// <summary>Uses the canonical ZIP scanner and matches the selected member without re-parsing a second BCL entry index.</summary>
        private bool TryGetZipCoordinates(ZipArchiveEntry entry, CancellationToken cancellationToken,
            out OfficeArchiveSafety.ZipCentralDirectoryRecord selected) {
            selected = default;
            if (!_stream.CanSeek) return false;
            foreach (char character in entry.FullName) if (character > 127) return false;
            if (!_checkedZipCoordinates) {
                long position = _stream.Position;
                try {
                    _stream.Position = 0;
                    var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
                    var scan = OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(
                        _stream, _stream.Length, _archiveEntryCount, records, cancellationToken);
                    if (scan.IsValid && !scan.LimitExceeded && scan.EntryCount == _archiveEntryCount)
                        _zipCoordinates = records;
                    _checkedZipCoordinates = true;
                } catch (InvalidDataException) {
                    _checkedZipCoordinates = true;
                } finally {
                    _stream.Position = position;
                }
            }
            if (_zipCoordinates == null) return false;
            byte[] name = Encoding.ASCII.GetBytes(entry.FullName);
            bool found = false;
            foreach (var record in _zipCoordinates) {
                cancellationToken.ThrowIfCancellationRequested();
                if (record.NameBytes.Length != name.Length) continue;
                bool same = true;
                for (int index = 0; index < name.Length; index++) {
                    if (record.NameBytes[index] == name[index]) continue;
                    same = false;
                    break;
                }
                if (!same) continue;
                if (found) return false;
                selected = record;
                found = true;
            }
            return found && OfficeArchiveSafety.HasSupportedZipFlags(selected.Flags, selected.CompressionMethod)
                && selected.CompressedSize == entry.CompressedLength && selected.UncompressedSize == entry.Length;
        }
    }
}
