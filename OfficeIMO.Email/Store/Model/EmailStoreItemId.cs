namespace OfficeIMO.Email.Store;

/// <summary>Format-neutral stable identifier for one item within an open email store.</summary>
public readonly struct EmailStoreItemId : IEquatable<EmailStoreItemId>, IComparable<EmailStoreItemId> {
    private readonly string? _value;

    /// <summary>Creates an item identifier from the exact source value.</summary>
    public EmailStoreItemId(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An item identifier cannot be empty.", nameof(value));
        _value = value;
    }

    /// <summary>Exact source identifier.</summary>
    public string Value => _value ?? throw new InvalidOperationException("The default item identifier is not valid.");

    /// <summary>Whether this is the uninitialized default struct value.</summary>
    public bool IsEmpty => _value == null;

    /// <summary>Parses a non-empty item identifier.</summary>
    public static EmailStoreItemId Parse(string value) => new EmailStoreItemId(value);

    /// <summary>Attempts to parse a non-empty item identifier.</summary>
    public static bool TryParse(string? value, out EmailStoreItemId id) {
        if (string.IsNullOrWhiteSpace(value)) {
            id = default;
            return false;
        }
        id = new EmailStoreItemId(value!);
        return true;
    }

    /// <inheritdoc />
    public bool Equals(EmailStoreItemId other) => StringComparer.Ordinal.Equals(_value, other._value);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is EmailStoreItemId other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => _value == null ? 0 : StringComparer.Ordinal.GetHashCode(_value);

    /// <inheritdoc />
    public int CompareTo(EmailStoreItemId other) => StringComparer.Ordinal.Compare(_value, other._value);

    /// <inheritdoc />
    public override string ToString() => _value ?? string.Empty;

    /// <summary>Tests two identifiers for ordinal equality.</summary>
    public static bool operator ==(EmailStoreItemId left, EmailStoreItemId right) => left.Equals(right);

    /// <summary>Tests two identifiers for ordinal inequality.</summary>
    public static bool operator !=(EmailStoreItemId left, EmailStoreItemId right) => !left.Equals(right);
}
