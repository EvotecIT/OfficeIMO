namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private static bool HasMergeSectionContent(RtfDocument document) => document.Blocks.Count > 0 ||
        document.Sections.Count > 0 || document.HeaderFooters.Count > 0 || document.PageSetup.HasAnyValue;

    private void AppendMergedSections(RtfDocument source, Dictionary<int, int> fonts, Dictionary<int, int> colors,
        Dictionary<int, int> revisionAuthors, ISet<RtfNote> remappedNotes, MergeResourceMap bindings) {
        if (!HasMergeSectionContent(source)) return;
        if (_sections.Count == 0 && _blocks.Count > 0) {
            var initial = new RtfSection(this);
            foreach (IRtfBlock block in _blocks) initial.AddParsedBlock(block);
            _sections.Add(initial);
        }
        IReadOnlyList<RtfHeaderFooter> firstHeaders = source.Sections.Count == 0
            ? source.HeaderFooters : source.GetEffectiveHeaderFooters(source.Sections[0]);
        var sections = source.Sections.Count == 0 ? new List<RtfSection>() : source.Sections.ToList();
        if (sections.Count == 0) {
            var section = new RtfSection(source);
            foreach (IRtfBlock block in source.Blocks) section.AddParsedBlock(block);
            sections.Add(section);
        }
        // An independent document must not inherit the preceding destination's header stories.
        foreach (RtfHeaderFooterKind kind in Enum.GetValues(typeof(RtfHeaderFooterKind))) {
            RtfHeaderFooter? declaration = firstHeaders.LastOrDefault(item => item.Kind == kind);
            if (declaration == null) {
                declaration = new RtfHeaderFooter(kind);
                RtfHeaderFooterKind? fallbackKind = kind is RtfHeaderFooterKind.LeftHeader or RtfHeaderFooterKind.RightHeader ? RtfHeaderFooterKind.Header
                    : kind is RtfHeaderFooterKind.LeftFooter or RtfHeaderFooterKind.RightFooter ? RtfHeaderFooterKind.Footer : (RtfHeaderFooterKind?)null;
                RtfHeaderFooter? fallback = fallbackKind.HasValue ? firstHeaders.LastOrDefault(item => item.Kind == fallbackKind.Value) : null;
                if (fallback != null) foreach (RtfParagraph paragraph in fallback.Paragraphs) declaration.AddParsedParagraph(paragraph);
            }
            sections[0].AddParsedHeaderFooter(declaration);
        }
        foreach (RtfSection section in sections) {
            section.PrepareMergedSection(source, this, colors);
            foreach (RtfHeaderFooter declaration in section.HeaderFooters) {
                foreach (RtfParagraph paragraph in declaration.Paragraphs)
                    RemapMergedParagraph(paragraph, fonts, colors, revisionAuthors, remappedNotes, bindings);
                if (!_headerFooters.Contains(declaration)) _headerFooters.Add(declaration);
            }
            AddParsedSection(section);
        }
        foreach (IRtfBlock block in source.Blocks) AddParsedBlock(block);
    }
}

public sealed partial class RtfSection {
    internal void PrepareMergedSection(RtfDocument source, RtfDocument destination, IReadOnlyDictionary<int, int> colors) {
        if (!ReferenceEquals(_document, source)) throw new InvalidOperationException("Only an independently cloned source section may be transferred.");
        PageSetup = PageSetup.WithFallback(source.PageSetup);
        // Materialize RTF's document defaults so destination-wide page setup cannot leak into the appended sections.
        PageSetup.PaperWidthTwips ??= 12240;
        PageSetup.PaperHeightTwips ??= 15840;
        PageSetup.MarginLeftTwips ??= 1800;
        PageSetup.MarginRightTwips ??= 1800;
        PageSetup.MarginTopTwips ??= 1440;
        PageSetup.MarginBottomTwips ??= 1440;
        PageSetup.GutterWidthTwips ??= 0;
        PageSetup.HeaderDistanceTwips ??= 720;
        PageSetup.FooterDistanceTwips ??= 720;
        PageSetup.DirectLandscape ??= false;
        PageSetup.DirectDifferentFirstPageHeaderFooter ??= false;
        PageSetup.DirectRtlGutter ??= false;
        foreach (RtfPageBorder border in new[] { PageSetup.PageBorders.Top, PageSetup.PageBorders.Bottom, PageSetup.PageBorders.Left, PageSetup.PageBorders.Right })
            if (border.ColorIndex.HasValue && colors.TryGetValue(border.ColorIndex.Value, out int color)) border.ColorIndex = color;
        NoteSettings.InheritMergeDefaults(source.NoteSettings);
        _document = destination;
    }
}

public sealed partial class RtfNoteSettings {
    internal void InheritMergeDefaults(RtfNoteSettings source) {
        FootnoteStartNumber ??= source.FootnoteStartNumber;
        FootnoteRestart ??= source.FootnoteRestart;
        FootnoteNumberFormat ??= source.FootnoteNumberFormat;
        FootnotePlacement ??= source.FootnotePlacement;
        EndnoteStartNumber ??= source.EndnoteStartNumber;
        EndnoteRestart ??= source.EndnoteRestart;
        EndnoteNumberFormat ??= source.EndnoteNumberFormat;
        EndnotePlacement ??= source.EndnotePlacement;
    }
}
