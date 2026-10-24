namespace OfficeIMO.Rtf;

public sealed partial class RtfDocumentSettings {
    private RtfNoteMode? _noteMode;

    /// <summary>Gets or sets the declared document note mode. Null retains the applicable parsed or authored default.</summary>
    /// <remarks>Parsed RTF without a declaration uses the legacy mode. Freshly authored documents use the modern mixed-note mode.</remarks>
    public RtfNoteMode? NoteMode {
        get => _noteMode;
        set {
            if (value.HasValue && value.Value != RtfNoteMode.Footnotes && value.Value != RtfNoteMode.Endnotes &&
                value.Value != RtfNoteMode.FootnotesAndEndnotes)
                throw new ArgumentOutOfRangeException(nameof(value), "RTF note mode must be 0, 1, or 2.");
            _noteMode = value;
        }
    }

    // This semantic provenance survives cloning and edits independently of retained source bytes.
    internal bool HasParsedNoteDefaults { get; set; }

    internal RtfNoteMode EffectiveNoteMode => NoteMode ??
        (HasParsedNoteDefaults ? RtfNoteMode.Footnotes : RtfNoteMode.FootnotesAndEndnotes);
}
