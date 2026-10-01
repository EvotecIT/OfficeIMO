namespace OfficeIMO.Rtf;

/// <summary>Controls whether appended content uses the destination layout or retains its source sections.</summary>
public sealed class RtfDocumentMergeOptions {
    /// <summary>Imports source sections, page setup, and header/footer stories. The default appends content into the destination layout.</summary>
    public bool PreserveSections { get; set; }
}
