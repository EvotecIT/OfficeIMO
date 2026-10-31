using System.Collections.Generic;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Keep references among their real run children. Text projection continues to use
    // the ordinary projectors, including casing, tabs, spacing, fonts and hyperlinks.
    private static bool HasNativeNoteContent(WordParagraph run, NativeNoteNumbering? notes,
        NativeDocumentDefaults defaults) => run.EnumerateEffectiveRunContent().Any(child =>
            GetNativeNoteReferenceKey(child) is NativeNoteKey key && notes?.TryGetValue(key, out _) == true ||
            (child is W.FootnoteReferenceMark or W.EndnoteReferenceMark) && defaults.NoteBodyLabel != null ||
            (child is W.SeparatorMark or W.ContinuationSeparatorMark) && defaults.NoteContentSection != null);

    private static bool TryAddNativeNoteRun(PdfCore.PdfParagraphBuilder builder, WordParagraph run,
        WordParagraph paragraph, NativeNoteNumbering notes, IReadOnlyList<WordTabStop> tabStops,
        ref int tabIndex, WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts,
        ISet<int> emittedNoteTokens) {
        if (!HasNativeNoteContent(run, notes, defaults)) return false;
        foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
            if (child is W.RunProperties) continue;
            if (TryCreateNativeNoteMarker(child, run, paragraph, notes, defaults, fonts,
                NativeTableRunStyleDefaults.Empty, false, out PdfCore.PdfTextRun? marker, out int? token)) {
                if (marker != null) builder.Runs(new[] { marker });
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
        foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
            if (child is W.RunProperties) continue;
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

    private static bool TryCreateNativeNoteMarker(OpenXmlElement child, WordParagraph source,
        WordParagraph paragraph, NativeNoteNumbering? notes, NativeDocumentDefaults defaults,
        NativeFontMap? fonts, NativeTableRunStyleDefaults table, bool configuredTypography,
        out PdfCore.PdfTextRun? marker, out int? noteToken) {
        marker = null;
        noteToken = null;
        if ((child is W.SeparatorMark or W.ContinuationSeparatorMark) && defaults.NoteContentSection != null) {
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
        } else if ((child is W.FootnoteReferenceMark or W.EndnoteReferenceMark) && defaults.NoteBodyLabel != null) {
            label = defaults.NoteBodyLabel;
        } else return false;
        NativeResolvedTextStyle style = ResolveNativeTextRunStyle(source, paragraph, table, defaults, fonts);
        marker = style.ApplyTo(new PdfCore.PdfTextRun(label, bold: style.Bold,
            underline: style.Underline, italic: style.Italic, strike: style.Strike,
            color: style.Color, fontSize: ((double?)(style.FontSize ?? (configuredTypography ? null : defaults.FontSize)) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null),
            font: style.Font, fontFamily: style.FontFamily, baseline: style.Baseline,
            backgroundColor: style.BackgroundColor, underlineStyle: style.UnderlineStyle,
            strikeStyle: style.StrikeStyle, linkDestinationName: destination));
        if (definition != null) marker = marker.WithFootnoteReference(definition);
        return true;
    }

    private static NativeNoteKey? GetNativeNoteReferenceKey(OpenXmlElement child) => child switch {
        W.FootnoteReference footnote when footnote.Id?.Value is long footnoteId => GetNativeFootnoteKey(footnoteId),
        W.EndnoteReference endnote when endnote.Id?.Value is long endnoteId => GetNativeEndnoteKey(endnoteId),
        _ => null
    };

    private static string GetNativePageNoteDestination(NativeNoteKey key) =>
        (key.Kind == NativeNoteKind.Endnote ? "OfficeIMO.Endnote." : "OfficeIMO.Footnote.") +
        key.Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
