namespace OfficeIMO.Pdf;

/// <summary>Chooses the expandable boundaries of a justified rich paragraph.</summary>
public enum PdfJustificationMode {
    /// <summary>Share extra line width between expandable word separators.</summary>
    Words,
    /// <summary>
    /// Share extra line width between shaped grapheme clusters, including ordinary spaces.
    /// Lines containing a tab preserve its fixed anchor and expand only word separators in the final tab field.
    /// </summary>
    Characters
}
