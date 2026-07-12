namespace OfficeIMO.Email;

/// <summary>Represents one ordered message header field.</summary>
public sealed class EmailHeader {
    /// <summary>Creates a header field.</summary>
    public EmailHeader(string name, string value, string? rawValue = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Header name is required.", nameof(name));
        Name = name;
        Value = value ?? string.Empty;
        RawValue = rawValue;
    }

    /// <summary>Header field name.</summary>
    public string Name { get; set; }

    /// <summary>Unfolded and decoded header value.</summary>
    public string Value { get; set; }

    /// <summary>Original unfolded value before encoded-word decoding.</summary>
    public string? RawValue { get; set; }
}
