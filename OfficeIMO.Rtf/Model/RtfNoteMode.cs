namespace OfficeIMO.Rtf;

/// <summary>Document note-mode declarations represented by the RTF <c>\fet</c> control.</summary>
/// <remarks>The declaration and an individual note's raw kind are separate from its native placement.</remarks>
public enum RtfNoteMode {
    /// <summary>The legacy footnote policy, represented by <c>\fet0</c>.</summary>
    Footnotes = 0,
    /// <summary>The endnote mode declaration, represented by <c>\fet1</c>.</summary>
    Endnotes = 1,
    /// <summary>The modern mixed-note declaration, represented by <c>\fet2</c>.</summary>
    FootnotesAndEndnotes = 2
}
