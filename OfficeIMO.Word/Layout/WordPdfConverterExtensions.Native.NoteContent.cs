using System.Collections.Generic;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Keep references among their real run children. Text projection continues to use
    // the ordinary projectors, including casing, tabs, spacing, fonts and hyperlinks.
    private static bool HasNativePageNoteContent(WordParagraph run, NativeNoteNumbering? notes,
        NativeDocumentDefaults defaults) => run.EnumerateEffectiveRunContent().Any(child =>
            child is W.FootnoteReference reference && reference.Id?.Value is long id &&
                notes?.TryGetPageFootnote(GetNativeFootnoteKey(id), out _) == true ||
            child is W.FootnoteReferenceMark && defaults.NoteBodyLabel != null);

    private static bool TryAddNativePageNoteRun(PdfCore.PdfParagraphBuilder builder, WordParagraph run,
        WordParagraph paragraph, NativeNoteNumbering notes, IReadOnlyList<WordTabStop> tabStops,
        ref int tabIndex, WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fonts) {
        if (!HasNativePageNoteContent(run, notes, defaults)) return false;
        foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
            if (child is W.RunProperties) continue;
            if (TryCreateNativePageNoteMarker(child, run, paragraph, notes, defaults, fonts,
                NativeTableRunStyleDefaults.Empty, false, out PdfCore.PdfTextRun? marker)) {
                builder.Runs(new[] { marker! });
            } else {
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                AddNativeRun(builder, view, paragraph, tabStops, ref tabIndex, options, defaults, fonts);
            }
        }
        return true;
    }

    private static bool TryAddNativePageNoteCellRun(List<PdfCore.PdfTextRun> result, WordParagraph run,
        NativeNoteNumbering? notes, NativeTableStyleDefaults table, NativeDocumentDefaults defaults,
        NativeFontMap? fonts, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
        if (!HasNativePageNoteContent(run, notes, defaults)) return false;
        foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
            if (child is W.RunProperties) continue;
            if (TryCreateNativePageNoteMarker(child, run, run, notes, defaults, fonts,
                table.RunStyle, table.UseConfiguredTypography, out PdfCore.PdfTextRun? marker)) {
                result.Add(marker!);
            } else {
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                AddNativeCellRun(result, view, table, defaults, fonts, tabStops, ref tabIndex);
            }
        }
        return true;
    }

    private static bool TryCreateNativePageNoteMarker(OpenXmlElement child, WordParagraph source,
        WordParagraph paragraph, NativeNoteNumbering? notes, NativeDocumentDefaults defaults,
        NativeFontMap? fonts, NativeTableRunStyleDefaults table, bool configuredTypography,
        out PdfCore.PdfTextRun? marker) {
        marker = null;
        string? label;
        string? destination = null;
        PdfCore.PdfFootnoteDefinition? definition = null;
        if (child is W.FootnoteReference reference && reference.Id?.Value is long id &&
            notes?.TryGetPageFootnote(GetNativeFootnoteKey(id), out definition) == true &&
            notes.TryGetValue(GetNativeFootnoteKey(id), out int token)) {
            label = notes.GetLabel(token);
            destination = string.IsNullOrEmpty(label) ? null : GetNativePageNoteDestination(GetNativeFootnoteKey(id));
        } else if (child is W.FootnoteReferenceMark && defaults.NoteBodyLabel != null) {
            label = defaults.NoteBodyLabel;
        } else return false;
        NativeResolvedTextStyle style = ResolveNativeTextRunStyle(source, paragraph, table, defaults, fonts);
        marker = style.ApplyTo(new PdfCore.PdfTextRun(label, bold: style.Bold,
            underline: style.Underline, italic: style.Italic, strike: style.Strike,
            color: style.Color, fontSize: style.FontSize ?? (configuredTypography ? null : defaults.FontSize),
            font: style.Font, fontFamily: style.FontFamily, baseline: PdfCore.PdfTextBaseline.Superscript,
            backgroundColor: style.BackgroundColor, underlineStyle: style.UnderlineStyle,
            strikeStyle: style.StrikeStyle, linkDestinationName: destination));
        if (definition != null) marker = marker.WithFootnoteReference(definition);
        return true;
    }

    private static string GetNativePageNoteDestination(NativeNoteKey key) =>
        "OfficeIMO.Footnote." + key.Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
