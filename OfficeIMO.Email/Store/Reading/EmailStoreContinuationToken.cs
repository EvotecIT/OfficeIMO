namespace OfficeIMO.Email.Store;

/// <summary>Opaque, portable keyset continuation for one exact Store query plan.</summary>
public sealed class EmailStoreContinuationToken : IEquatable<EmailStoreContinuationToken> {
    private const int Magic = 0x314D494F;
    private const byte Version = 1;
    private const int MaxEncodedLength = 128 * 1024;
    private readonly IReadOnlyList<byte[]> _sortValues;

    private EmailStoreContinuationToken(string value, string querySignature, IReadOnlyList<byte[]> sortValues) {
        Value = value;
        QuerySignature = querySignature;
        _sortValues = sortValues;
    }

    /// <summary>URL/file-safe Base64 value suitable for persistence or API transport.</summary>
    public string Value { get; }

    internal string QuerySignature { get; }
    internal IReadOnlyList<byte[]> SortValues => _sortValues;

    /// <summary>Parses and structurally validates a continuation token.</summary>
    public static EmailStoreContinuationToken Parse(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A continuation token cannot be empty.", nameof(value));
        if (value.Length > MaxEncodedLength) throw new InvalidDataException("The continuation token is too large.");
        byte[] bytes;
        try {
            string normalized = value.Replace('-', '+').Replace('_', '/');
            switch (normalized.Length % 4) {
                case 2: normalized += "=="; break;
                case 3: normalized += "="; break;
            }
            bytes = Convert.FromBase64String(normalized);
        } catch (FormatException exception) {
            throw new InvalidDataException("The continuation token is not valid Base64.", exception);
        }
        using (var stream = new MemoryStream(bytes, writable: false))
        using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true)) {
            if (reader.ReadInt32() != Magic || reader.ReadByte() != Version) {
                throw new InvalidDataException("The continuation token format or version is unsupported.");
            }
            string signature = EmailStoreScalarCodec.ReadString(reader, 65_536);
            int count = reader.ReadInt32();
            if (count <= 0 || count > 64) throw new InvalidDataException("The continuation token sort-key count is invalid.");
            var values = new List<byte[]>(count);
            for (int index = 0; index < count; index++) {
                int length = reader.ReadInt32();
                if (length <= 0 || length > 65_536 || length > stream.Length - stream.Position) {
                    throw new InvalidDataException("The continuation token sort-key length is invalid.");
                }
                byte[] scalar = reader.ReadBytes(length);
                if (scalar.Length != length) throw new EndOfStreamException();
                values.Add(scalar);
            }
            if (stream.Position != stream.Length) throw new InvalidDataException("The continuation token contains trailing data.");
            return new EmailStoreContinuationToken(value, signature, values.AsReadOnly());
        }
    }

    /// <summary>Attempts to parse and structurally validate a continuation token.</summary>
    public static bool TryParse(string? value, out EmailStoreContinuationToken? token) {
        try {
            token = value == null ? null : Parse(value);
            return token != null;
        } catch (ArgumentException) {
            token = null;
            return false;
        } catch (InvalidDataException) {
            token = null;
            return false;
        } catch (EndOfStreamException) {
            token = null;
            return false;
        }
    }

    internal static EmailStoreContinuationToken Create(EmailStoreTableQuery query, EmailStoreQueryRow row) {
        var values = query.EffectiveSorts
            .Select(sort => EmailStoreScalarCodec.Serialize(sort.Field.Read(row)))
            .ToArray();
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(Magic);
            writer.Write(Version);
            EmailStoreScalarCodec.WriteString(writer, query.Signature);
            writer.Write(values.Length);
            foreach (byte[] value in values) {
                writer.Write(value.Length);
                writer.Write(value);
            }
            writer.Flush();
            string encoded = Convert.ToBase64String(stream.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
            return new EmailStoreContinuationToken(encoded, query.Signature, Array.AsReadOnly(values));
        }
    }

    internal IReadOnlyList<object?> DecodeValues(EmailStoreTableQuery query) {
        if (!StringComparer.Ordinal.Equals(QuerySignature, query.Signature)) {
            throw new ArgumentException("The continuation token belongs to a different Store query scope, filter, or ordering.", nameof(query));
        }
        if (_sortValues.Count != query.EffectiveSorts.Count) {
            throw new InvalidDataException("The continuation token does not contain the expected sort keys.");
        }
        var values = new object?[_sortValues.Count];
        for (int index = 0; index < values.Length; index++) {
            values[index] = EmailStoreScalarCodec.Deserialize(
                _sortValues[index], query.EffectiveSorts[index].Field.ValueType);
        }
        return Array.AsReadOnly(values);
    }

    /// <inheritdoc />
    public bool Equals(EmailStoreContinuationToken? other) => other != null && StringComparer.Ordinal.Equals(Value, other.Value);

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as EmailStoreContinuationToken);

    /// <inheritdoc />
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Value);

    /// <inheritdoc />
    public override string ToString() => Value;
}
