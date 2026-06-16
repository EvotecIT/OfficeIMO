namespace OfficeIMO.Rtf;

/// <summary>
/// Document variable from the RTF <c>\docvar</c> destination.
/// </summary>
public sealed class RtfDocumentVariable {
    /// <summary>Creates a document variable.</summary>
    public RtfDocumentVariable(string name, string value) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Document variable name cannot be empty.", nameof(name));
        Name = name;
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Document variable name.</summary>
    public string Name { get; set; }

    /// <summary>Document variable value.</summary>
    public string Value { get; set; }
}
