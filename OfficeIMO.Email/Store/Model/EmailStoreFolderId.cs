namespace OfficeIMO.Email.Store;

/// <summary>Format-neutral stable identifier for one folder within an open email store.</summary>
public readonly struct EmailStoreFolderId : IEquatable<EmailStoreFolderId>, IComparable<EmailStoreFolderId> {
    private readonly string? _value;

    /// <summary>Creates a folder identifier from the exact source value.</summary>
    public EmailStoreFolderId(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A folder identifier cannot be empty.", nameof(value));
        _value = value;
    }

    /// <summary>Exact source identifier.</summary>
    public string Value => _value ?? throw new InvalidOperationException("The default folder identifier is not valid.");

    /// <summary>Whether this is the uninitialized default struct value.</summary>
    public bool IsEmpty => _value == null;

    /// <summary>Parses a non-empty folder identifier.</summary>
    public static EmailStoreFolderId Parse(string value) => new EmailStoreFolderId(value);

    /// <summary>Attempts to parse a non-empty folder identifier.</summary>
    public static bool TryParse(string? value, out EmailStoreFolderId id) {
        if (string.IsNullOrWhiteSpace(value)) {
            id = default;
            return false;
        }
        id = new EmailStoreFolderId(value!);
        return true;
    }

    /// <inheritdoc />
    public bool Equals(EmailStoreFolderId other) => StringComparer.Ordinal.Equals(_value, other._value);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is EmailStoreFolderId other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => _value == null ? 0 : StringComparer.Ordinal.GetHashCode(_value);

    /// <inheritdoc />
    public int CompareTo(EmailStoreFolderId other) => StringComparer.Ordinal.Compare(_value, other._value);

    /// <inheritdoc />
    public override string ToString() => _value ?? string.Empty;

    /// <summary>Tests two identifiers for ordinal equality.</summary>
    public static bool operator ==(EmailStoreFolderId left, EmailStoreFolderId right) => left.Equals(right);

    /// <summary>Tests two identifiers for ordinal inequality.</summary>
    public static bool operator !=(EmailStoreFolderId left, EmailStoreFolderId right) => !left.Equals(right);
}
