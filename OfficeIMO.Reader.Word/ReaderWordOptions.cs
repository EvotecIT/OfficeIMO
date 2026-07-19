namespace OfficeIMO.Reader.Word;

/// <summary>Controls Word document projection into Reader chunks.</summary>
public sealed class ReaderWordOptions {
    /// <summary>Includes footnotes in the Markdown projection.</summary>
    public bool IncludeFootnotes { get; set; } = true;
}
