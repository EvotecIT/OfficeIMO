using OfficeIMO.Core.Internal;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    /// <summary>Embedded attachment metadata with lazy, bounded native payload decoding. Files are never opened or executed.</summary>
    public sealed class AccessAttachment {
        private readonly AccessNativeTable _table;
        private readonly AccessNativeRow _row;
        private readonly CancellationToken _cancellation;
        internal AccessAttachment(AccessNativeTable table, AccessNativeRow row, CancellationToken cancellation) {
            _table = table; _row = row; _cancellation = cancellation;
            FileName = Value("FileName") as string; FileType = Value("FileType") as string; FileUrl = Value("FileURL") as string;
            TimeStamp = Value("FileTimeStamp") as DateTime?; Flags = Value("FileFlags") as int?;
        }
        private object? Value(string name) {
            _table.Database.Document.EnsureNotDisposed(); _cancellation.ThrowIfCancellationRequested();
            int ordinal = _table.Columns.FindIndex(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, name));
            if (ordinal < 0) throw new InvalidDataException("Native Access attachment is missing a required metadata field.");
            return _row.Value(ordinal, _cancellation);
        }
        /// <summary>Stored file name; it is never used as an output path.</summary>
        public string? FileName { get; }
        /// <summary>Stored file extension/type label.</summary>
        public string? FileType { get; }
        /// <summary>Stored URL metadata; loading never follows it.</summary>
        public string? FileUrl { get; }
        /// <summary>Optional native file timestamp, without a guessed time zone.</summary>
        public DateTime? TimeStamp { get; }
        /// <summary>Persisted attachment flags.</summary>
        public int? Flags { get; }
        /// <summary>Returns a defensive copy of the exact native attachment wrapper, including compression and metadata.</summary>
        public byte[] GetEncodedBytes() => (byte[])((byte[]?)Value("FileData") ?? Array.Empty<byte>()).Clone();
        /// <summary>Decodes the bounded embedded file bytes through Core's checksum-validating compression owner.</summary>
        public byte[] GetBytes(CancellationToken cancellationToken = default) {
            using CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(_cancellation, cancellationToken);
            linked.Token.ThrowIfCancellationRequested(); byte[] encoded = GetEncodedBytes();
            int kind = I32(encoded, 0), length = I32(encoded, 4);
            if (length < 12 || length > _table.Database.MaxValueBytes) throw new InvalidDataException("Native Access attachment exceeds its decoded value limit or has an invalid content length.");
            byte[] content;
            if (kind == 0) { if (encoded.Length - 8 != length) throw new InvalidDataException("Native Access raw attachment length is inconsistent."); content = Slice(encoded, 8, length).ToArray(); }
            else if (kind == 1) content = OfficeZlibCodec.Decompress(Slice(encoded, 8, encoded.Length - 8).ToArray(), _table.Database.MaxValueBytes, length, linked.Token);
            else throw new NotSupportedException("Native Access attachment compression kind is unqualified; GetEncodedBytes retains its exact representation.");
            int headerLength = I32(content, 0), characters = I32(content, 8);
            if (headerLength < 12 || headerLength > length || characters < 1 || characters > (headerLength - 12) / 2 || I32(content, 4) != 1) throw new InvalidDataException("Native Access attachment content header is malformed.");
            linked.Token.ThrowIfCancellationRequested(); return Slice(content, headerLength, length - headerLength).ToArray();
        }
        /// <summary>Opens a caller-owned read-only stream over decoded file bytes. No filesystem path is resolved.</summary>
        public Stream OpenRead(CancellationToken cancellationToken = default) =>
            new OfficeDocumentReadStream(new MemoryStream(GetBytes(cancellationToken), writable: false),
                () => { _table.Database.Document.EnsureNotDisposed(); _cancellation.ThrowIfCancellationRequested(); }, cancellationToken);
    }
}
