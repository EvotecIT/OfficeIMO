namespace OfficeIMO.Email.AddressBook;

/// <summary>Portable source- and query-bound exact-position checkpoint for OAB search.</summary>
public sealed class OfflineAddressBookSearchCheckpoint {
    private const int Magic = 0x4241494F;
    private const byte Version = 1;

    private OfflineAddressBookSearchCheckpoint(string value, string addressListId, int addressListIndex,
        long entryIndex, long recordOffset, string sourceFingerprint, string querySignature) {
        Value = value; AddressListId = addressListId; AddressListIndex = addressListIndex;
        EntryIndex = entryIndex; RecordOffset = recordOffset; SourceFingerprint = sourceFingerprint;
        QuerySignature = querySignature;
    }

    /// <summary>URL/file-safe Base64 checkpoint suitable for persistence across processes.</summary>
    public string Value { get; }
    /// <summary>Address-list identifier.</summary>
    public string AddressListId { get; }
    /// <summary>Zero-based address-list index.</summary>
    public int AddressListIndex { get; }
    /// <summary>Zero-based index of the next record to scan.</summary>
    public long EntryIndex { get; }
    /// <summary>Exact next-record offset relative to its Full Details component.</summary>
    public long RecordOffset { get; }
    /// <summary>SHA-256 identity over the complete OAB source set.</summary>
    public string SourceFingerprint { get; }
    internal string QuerySignature { get; }

    /// <summary>Parses and structurally validates a durable checkpoint.</summary>
    public static OfflineAddressBookSearchCheckpoint Parse(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A checkpoint cannot be empty.", nameof(value));
        if (value.Length > 4096) throw new InvalidDataException("The checkpoint is too large.");
        try {
            string normalized = value.Replace('-', '+').Replace('_', '/');
            if (normalized.Length % 4 == 2) normalized += "==";
            else if (normalized.Length % 4 == 3) normalized += "=";
            byte[] bytes = Convert.FromBase64String(normalized);
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
            if (reader.ReadInt32() != Magic || reader.ReadByte() != Version) throw new InvalidDataException("The checkpoint version is unsupported.");
            string id = ReadString(reader, 1024); int index = reader.ReadInt32();
            long entry = reader.ReadInt64(); long offset = reader.ReadInt64();
            string source = ReadString(reader, 256); string query = ReadString(reader, 256);
            if (index < 0 || entry < 0 || offset < 0 || source.Length != 64 || query.Length != 64 || stream.Position != stream.Length) {
                throw new InvalidDataException("The checkpoint payload is invalid.");
            }
            return new OfflineAddressBookSearchCheckpoint(value, id, index, entry, offset, source, query);
        } catch (FormatException exception) { throw new InvalidDataException("The checkpoint is not valid Base64.", exception); }
    }

    /// <summary>Attempts to parse a durable checkpoint.</summary>
    public static bool TryParse(string? value, out OfflineAddressBookSearchCheckpoint? checkpoint) {
        try { checkpoint = value == null ? null : Parse(value); return checkpoint != null; }
        catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException || exception is EndOfStreamException) { checkpoint = null; return false; }
    }

    internal static OfflineAddressBookSearchCheckpoint Create(string addressListId, int addressListIndex,
        long entryIndex, long recordOffset, string sourceFingerprint, string querySignature) {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(Magic); writer.Write(Version); WriteString(writer, addressListId);
            writer.Write(addressListIndex); writer.Write(entryIndex); writer.Write(recordOffset);
            WriteString(writer, sourceFingerprint); WriteString(writer, querySignature);
        }
        string value = Convert.ToBase64String(stream.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        return new OfflineAddressBookSearchCheckpoint(value, addressListId, addressListIndex, entryIndex,
            recordOffset, sourceFingerprint, querySignature);
    }

    internal void Validate(string sourceFingerprint, string querySignature) {
        if (!StringComparer.Ordinal.Equals(SourceFingerprint, sourceFingerprint)) throw new ArgumentException("The OAB checkpoint belongs to a changed or different source.");
        if (!StringComparer.Ordinal.Equals(QuerySignature, querySignature)) throw new ArgumentException("The OAB checkpoint belongs to a different query.");
    }

    private static void WriteString(BinaryWriter writer, string value) { byte[] bytes = Encoding.UTF8.GetBytes(value); writer.Write(bytes.Length); writer.Write(bytes); }
    private static string ReadString(BinaryReader reader, int maximum) { int length = reader.ReadInt32(); if (length < 0 || length > maximum) throw new InvalidDataException("The checkpoint string length is invalid."); byte[] bytes = reader.ReadBytes(length); if (bytes.Length != length) throw new EndOfStreamException(); return Encoding.UTF8.GetString(bytes); }
    /// <inheritdoc />
    public override string ToString() => Value;
}
