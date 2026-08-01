namespace OfficeIMO.Email.Store;

/// <summary>Portable, source- and query-bound content-search resume position.</summary>
public sealed class EmailStoreContentSearchCheckpoint {
    private const int Magic = 0x4353494F;
    private const byte Version = 1;
    private const int MaxEncodedLength = 2048;

    private EmailStoreContentSearchCheckpoint(string value, long itemOffset, string sourceFingerprint,
        string querySignature) {
        Value = value;
        ItemOffset = itemOffset;
        SourceFingerprint = sourceFingerprint;
        QuerySignature = querySignature;
    }

    /// <summary>URL/file-safe Base64 value suitable for durable persistence.</summary>
    public string Value { get; }
    /// <summary>Number of item references already processed in the selected enumeration scope.</summary>
    public long ItemOffset { get; }
    /// <summary>SHA-256 fingerprint of the complete source.</summary>
    public string SourceFingerprint { get; }
    internal string QuerySignature { get; }

    /// <summary>Parses and structurally validates a durable checkpoint.</summary>
    public static EmailStoreContentSearchCheckpoint Parse(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A checkpoint cannot be empty.", nameof(value));
        if (value.Length > MaxEncodedLength) throw new InvalidDataException("The checkpoint is too large.");
        try {
            string normalized = value.Replace('-', '+').Replace('_', '/');
            if (normalized.Length % 4 == 2) normalized += "==";
            else if (normalized.Length % 4 == 3) normalized += "=";
            byte[] bytes = Convert.FromBase64String(normalized);
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
            if (reader.ReadInt32() != Magic || reader.ReadByte() != Version) throw new InvalidDataException("The checkpoint version is unsupported.");
            long offset = reader.ReadInt64();
            string source = EmailStoreScalarCodec.ReadString(reader, 256);
            string query = EmailStoreScalarCodec.ReadString(reader, 256);
            if (offset < 0 || offset > int.MaxValue || source.Length != 64 || query.Length != 64 || stream.Position != stream.Length) {
                throw new InvalidDataException("The checkpoint payload is invalid.");
            }
            return new EmailStoreContentSearchCheckpoint(value, offset, source, query);
        } catch (FormatException exception) {
            throw new InvalidDataException("The checkpoint is not valid Base64.", exception);
        }
    }

    /// <summary>Attempts to parse a durable checkpoint.</summary>
    public static bool TryParse(string? value, out EmailStoreContentSearchCheckpoint? checkpoint) {
        try { checkpoint = value == null ? null : Parse(value); return checkpoint != null; }
        catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException || exception is EndOfStreamException) {
            checkpoint = null; return false;
        }
    }

    internal static EmailStoreContentSearchCheckpoint Create(long itemOffset, string sourceFingerprint,
        string querySignature) {
        if (itemOffset < 0 || itemOffset > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(itemOffset));
        }
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(Magic); writer.Write(Version); writer.Write(itemOffset);
            EmailStoreScalarCodec.WriteString(writer, sourceFingerprint);
            EmailStoreScalarCodec.WriteString(writer, querySignature);
        }
        string value = Convert.ToBase64String(stream.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        return new EmailStoreContentSearchCheckpoint(value, itemOffset, sourceFingerprint, querySignature);
    }

    internal void Validate(string sourceFingerprint, string querySignature) {
        if (!StringComparer.Ordinal.Equals(SourceFingerprint, sourceFingerprint)) throw new ArgumentException("The content-search checkpoint belongs to a changed or different Store source.");
        if (!StringComparer.Ordinal.Equals(QuerySignature, querySignature)) throw new ArgumentException("The content-search checkpoint belongs to a different query.");
    }

    /// <inheritdoc />
    public override string ToString() => Value;
}
