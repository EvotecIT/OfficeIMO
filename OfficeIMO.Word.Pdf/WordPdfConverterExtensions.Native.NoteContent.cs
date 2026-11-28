using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Keep references among their real run children. Text projection continues to use
    // the ordinary projectors, including casing, tabs, spacing, fonts and hyperlinks.
    private static bool HasNativeNoteContent(WordParagraph run, NativeNoteNumbering? notes,
        NativeDocumentDefaults defaults) => run.EffectiveRunContentNodes().Any(child =>
            GetNativeNoteReferenceKey(child) is NativeNoteKey key && notes?.TryGetValue(key, out _) == true ||
            IsNativeNoteReferenceMark(child) && defaults.NoteBodyLabel != null ||
            IsNativeNoteSeparatorMark(child) && defaults.NoteContentSection != null);

    private static bool IsNativeNoteReferenceMark(OfficeXmlNode child) =>
        child.IsW("footnoteRef") || child.IsW("endnoteRef");

    private static bool IsNativeNoteSeparatorMark(OfficeXmlNode child) =>
        child.IsW("separator") || child.IsW("continuationSeparator");

    private static bool TryAddNativeNoteRun(PdfCore.PdfParagraph builder, WordParagraph run,
        WordParagraph paragraph, NativeNoteNumbering notes, IReadOnlyList<WordTabStop> tabStops,
        ref int tabIndex, WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts,
        ISet<int> emittedNoteTokens) {
        if (!HasNativeNoteContent(run, notes, defaults)) return false;
        foreach (OfficeXmlNode child in run.EffectiveRunContentNodes()) {
            if (child.IsW("rPr")) continue;
            if (TryCreateNativeNoteMarker(child, run, paragraph, notes, defaults, fonts,
                NativeTableRunStyleDefaults.Empty, false, out PdfCore.PdfTextRun? marker, out int? token)) {
                if (marker != null) builder.AddRuns(new[] { marker });
                if (token.HasValue) emittedNoteTokens.Add(token.Value);
            } else {
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                AddNativeRun(builder, view, paragraph, tabStops, ref tabIndex, options, defaults, fonts);
            }
        }
        return true;
    }

    private static bool TryAddNativeNoteCellRun(List<PdfCore.PdfTextRun> result, WordParagraph run,
        NativeNoteNumbering? notes, NativeTableStyleDefaults table, NativeDocumentDefaults defaults,
        NativeFontMap? fonts, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex, ISet<int> emittedNoteTokens) {
        if (!HasNativeNoteContent(run, notes, defaults)) return false;
        foreach (OfficeXmlNode child in run.EffectiveRunContentNodes()) {
            if (child.IsW("rPr")) continue;
            if (TryCreateNativeNoteMarker(child, run, run, notes, defaults, fonts,
                table.RunStyle, table.UseConfiguredTypography, out PdfCore.PdfTextRun? marker, out int? token)) {
                if (marker != null) result.Add(marker);
                if (token.HasValue) emittedNoteTokens.Add(token.Value);
            } else {
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                AddNativeCellRun(result, view, table, defaults, fonts, tabStops, ref tabIndex);
            }
        }
        return true;
    }

    private static bool TryCreateNativeNoteMarker(OfficeXmlNode child, WordParagraph source,
        WordParagraph paragraph, NativeNoteNumbering? notes, NativeDocumentDefaults defaults,
        NativeFontMap? fonts, NativeTableRunStyleDefaults table, bool configuredTypography,
        out PdfCore.PdfTextRun? marker, out int? noteToken) {
        marker = null;
        noteToken = null;
        if (IsNativeNoteSeparatorMark(child) && defaults.NoteContentSection != null) {
            marker = CreateNativeFootnoteSeparatorRule(child, source, defaults, fonts, table, notes);
            if (marker == null) notes?.ReportSeparatorMarkOmission();
            return true;
        }
        string? label;
        string? destination = null;
        PdfCore.PdfFootnoteDefinition? definition = null;
        NativeNoteKey? key = GetNativeNoteReferenceKey(child);
        if (key.HasValue && notes != null && notes.TryGetValue(key.Value, out int token)) {
            label = notes.GetLabel(token);
            noteToken = token;
            notes.TryGetPageFootnote(key.Value, out definition);
            destination = string.IsNullOrEmpty(label) ? null : GetNativePageNoteDestination(key.Value);
        } else if (IsNativeNoteReferenceMark(child) && defaults.NoteBodyLabel != null) {
            label = defaults.NoteBodyLabel;
        } else return false;
        NativeResolvedTextStyle style = ResolveNativeTextRunStyle(source, paragraph, table, defaults, fonts);
        marker = style.ApplyTo(new PdfCore.PdfTextRun(label, bold: style.Bold,
            underline: style.Underline, italic: style.Italic, strike: style.Strike,
            color: style.Color, fontSize: ((double?)(style.FontSize ?? (configuredTypography ? null : defaults.FontSize)) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null),
            standardFont: style.Font, fontFamily: style.FontFamily, baseline: style.Baseline,
            backgroundColor: style.BackgroundColor, underlineStyle: style.UnderlineStyle,
            strikeStyle: style.StrikeStyle, linkDestinationName: destination));
        if (definition != null) marker = marker.WithFootnoteReference(definition);
        return true;
    }

    private static NativeNoteKey? GetNativeNoteReferenceKey(OfficeXmlNode child) =>
        child.IsW("footnoteReference") && WordXml.ParseInt64(child.WAttr("id")) is long footnoteId ? GetNativeFootnoteKey(footnoteId) :
        child.IsW("endnoteReference") && WordXml.ParseInt64(child.WAttr("id")) is long endnoteId ? GetNativeEndnoteKey(endnoteId) :
        null;

    private static string GetNativePageNoteDestination(NativeNoteKey key) =>
        (key.Kind == NativeNoteKind.Endnote ? "OfficeIMO.Endnote." : "OfficeIMO.Footnote.") +
        key.Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
