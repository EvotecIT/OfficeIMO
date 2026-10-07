namespace OfficeIMO.Epub;

/// <summary>A scoped replacement or deletion guarded by the expected current element.</summary>
public sealed class EpubContentEdit {
    private readonly XElement _expected;
    private readonly XElement? _replacement;
    /// <summary>Creates an edit. XML is copied; a null replacement deletes the selected element.</summary>
    public EpubContentEdit(string manifestId, string elementId, XElement expected, XElement? replacement) {
        if (string.IsNullOrWhiteSpace(manifestId)) throw new ArgumentException("A manifest identifier is required.", nameof(manifestId));
        if (string.IsNullOrWhiteSpace(elementId)) throw new ArgumentException("An element identifier is required.", nameof(elementId));
        ManifestId = manifestId; ElementId = elementId;
        _expected = new XElement(expected ?? throw new ArgumentNullException(nameof(expected)));
        _replacement = replacement == null ? null : new XElement(replacement);
    }
    /// <summary>The existing content resource identifier.</summary>
    public string ManifestId { get; }
    /// <summary>The existing id or xml:id of the selected element.</summary>
    public string ElementId { get; }
    /// <summary>An independent copy of the expected current element.</summary>
    public XElement Expected => new XElement(_expected);
    /// <summary>An independent replacement copy, or null for deletion.</summary>
    public XElement? Replacement => _replacement == null ? null : new XElement(_replacement);
}
