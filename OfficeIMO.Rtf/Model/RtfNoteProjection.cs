namespace OfficeIMO.Rtf;

/// <summary>Resolves document/section note policy without changing source note collections.</summary>
internal sealed class RtfNoteProjection {
    private readonly RtfDocument _document;
    private readonly RtfSection? _section;
    private readonly RtfNoteMode _mode;
    private RtfNoteSettings? _projectedSettings;
    private readonly RtfNoteSettings _settings;

    private RtfNoteProjection(RtfDocument document, RtfSection? section, RtfNoteSettings settings) {
        _document = document;
        _section = section;
        _mode = document.Settings.EffectiveNoteMode;
        _settings = settings;
        FootnotesAreEndnotes = _mode == RtfNoteMode.Footnotes &&
            (!settings.FootnotePlacement.HasValue || settings.FootnotePlacement == RtfFootnotePlacement.SectionEnd ||
             settings.FootnotePlacement == RtfFootnotePlacement.DocumentEnd);
    }

    internal bool FootnotesAreEndnotes { get; }

    // Kind resolution is constant-time per reference. Inspect the owning scope only when applying group settings.
    internal RtfNoteSettings Settings => _projectedSettings ??= ProjectSettings();

    private RtfNoteSettings ProjectSettings() {
        RtfNoteSettings settings = _settings;
        IEnumerable<RtfNote> notes = _section == null
            ? RtfNoteReferenceCollector.Collect(_document).Concat(_document.Notes)
            : RtfNoteReferenceCollector.Collect(_section.Blocks);
        if (FootnotesAreEndnotes && notes.Any(note => note.Kind == RtfNoteKind.Footnote)) {
            // Legacy ordinary notes use the footnote group's numbering, with native endnote defaults.
            settings.EndnoteStartNumber = settings.FootnoteStartNumber;
            settings.EndnoteRestart = settings.FootnoteRestart;
            settings.EndnoteNumberFormat = settings.FootnoteNumberFormat ?? RtfNoteNumberFormat.LowerRoman;
            settings.EndnotePlacement = settings.FootnotePlacement == RtfFootnotePlacement.DocumentEnd
                ? RtfEndnotePlacement.DocumentEnd : RtfEndnotePlacement.SectionEnd;
            settings.FootnoteStartNumber = null;
            settings.FootnoteRestart = null;
            settings.FootnoteNumberFormat = null;
            settings.FootnotePlacement = null;
        } else if (_mode != RtfNoteMode.Footnotes &&
            (settings.FootnotePlacement == RtfFootnotePlacement.SectionEnd || settings.FootnotePlacement == RtfFootnotePlacement.DocumentEnd)) {
            // Modern modes keep ordinary footnotes at the page bottom despite legacy placement declarations.
            settings.FootnotePlacement = RtfFootnotePlacement.PageBottom;
        }
        return settings;
    }

    internal RtfNoteKind ProjectKind(RtfNoteKind sourceKind) =>
        FootnotesAreEndnotes && sourceKind == RtfNoteKind.Footnote ? RtfNoteKind.Endnote : sourceKind;

    internal static RtfNoteProjection Resolve(RtfDocument document, RtfSection? section = null) {
        RtfNoteSettings declared = document.NoteSettings;
        RtfNoteSettings? local = section?.NoteSettings;
        return new RtfNoteProjection(document, section, new RtfNoteSettings {
            FootnoteStartNumber = local?.FootnoteStartNumber ?? declared.FootnoteStartNumber,
            FootnoteRestart = local?.FootnoteRestart ?? declared.FootnoteRestart,
            FootnoteNumberFormat = local?.FootnoteNumberFormat ?? declared.FootnoteNumberFormat,
            FootnotePlacement = local?.FootnotePlacement ?? declared.FootnotePlacement,
            EndnoteStartNumber = local?.EndnoteStartNumber ?? declared.EndnoteStartNumber,
            EndnoteRestart = local?.EndnoteRestart ?? declared.EndnoteRestart,
            EndnoteNumberFormat = local?.EndnoteNumberFormat ?? declared.EndnoteNumberFormat,
            EndnotePlacement = local?.EndnotePlacement ?? declared.EndnotePlacement
        });
    }
}
