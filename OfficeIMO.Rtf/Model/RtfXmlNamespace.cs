namespace OfficeIMO.Rtf;

/// <summary>
/// XML namespace declaration from the RTF <c>{\*\xmlnstbl ...}</c> destination.
/// </summary>
public sealed class RtfXmlNamespace {
    /// <summary>
    /// Initializes an XML namespace declaration.
    /// </summary>
    public RtfXmlNamespace(int id, string uri) {
        if (id < 0) throw new ArgumentOutOfRangeException(nameof(id), "XML namespace id cannot be negative.");
        Id = id;
        Uri = uri ?? throw new ArgumentNullException(nameof(uri));
    }

    /// <summary>Namespace id represented by <c>\xmlnsN</c>.</summary>
    public int Id { get; }

    /// <summary>Namespace URI text.</summary>
    public string Uri { get; set; }
}
