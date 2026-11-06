namespace OfficeIMO.Rtf;

/// <summary>
/// Document variable from the RTF <c>\docvar</c> destination.
/// </summary>
public sealed partial class RtfDocumentVariable {
    /// <summary>Creates a document variable.</summary>
    public RtfDocumentVariable(string name, string value) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Name = name;
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Document variable name.</summary>
    public string Name { get; set; }

    /// <summary>Document variable value.</summary>
    public string Value { get; set; }
}
