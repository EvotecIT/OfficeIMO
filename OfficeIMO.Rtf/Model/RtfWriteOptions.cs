namespace OfficeIMO.Rtf;

/// <summary>
/// Options controlling RTF serialization.
/// </summary>
public sealed class RtfWriteOptions {
    /// <summary>Whether to emit a generator group from document metadata or the OfficeIMO default.</summary>
    public bool IncludeGenerator { get; set; } = true;

    /// <summary>Default font name used when the document has no fonts.</summary>
    public string DefaultFontName { get; set; } = "Calibri";
}
