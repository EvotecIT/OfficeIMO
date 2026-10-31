using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Generated reference text must retain its source formatting independently
        // of the minimum font used to lay out the surrounding mixed-size paragraph.
        private static IEnumerable<PdfCore.PdfTextRun> CreateNativeNoteReferenceRuns(
            WordParagraph paragraph, IReadOnlyList<int> numbers, NativeNoteNumbering numbersById,
            NativeDocumentDefaults defaults, NativeFontMap? fontMap,
            NativeTableRunStyleDefaults tableDefaults = default, bool useConfiguredTypography = false,
            ISet<int>? emittedNoteTokens = null) {
            var pageFootnoteTokens = new HashSet<int>();
            var sourceRuns = new Dictionary<int, W.Run>();
            int scannedRuns = 0;
            foreach (W.Run run in paragraph._paragraph.Descendants<W.Run>()) {
                if (++scannedRuns > 1_000_000) throw new InvalidDataException("The Word paragraph exceeds the native note-reference scan limit.");
                foreach (W.FootnoteReference reference in run.Elements<W.FootnoteReference>()) {
                    if (reference.Id?.Value is not long id) continue;
                    NativeNoteKey key = GetNativeFootnoteKey(id);
                    if (!numbersById.TryGetValue(key, out int token)) continue;
                    if (numbersById.TryGetPageFootnote(key, out _)) pageFootnoteTokens.Add(token);
                    else if (!sourceRuns.ContainsKey(token)) sourceRuns.Add(token, run);
                }
                foreach (W.EndnoteReference reference in run.Elements<W.EndnoteReference>()) {
                    if (reference.Id?.Value is not long id) continue;
                    NativeNoteKey key = GetNativeEndnoteKey(id);
                    if (!numbersById.TryGetPageFootnote(key, out _) &&
                        numbersById.TryGetValue(key, out int token) && !sourceRuns.ContainsKey(token)) sourceRuns.Add(token, run);
                }
            }
            foreach (int number in numbers) {
                if (emittedNoteTokens?.Contains(number) == true) continue;
                if (pageFootnoteTokens.Contains(number)) continue;
                WordParagraph? source = sourceRuns.TryGetValue(number, out W.Run? run)
                    ? new WordParagraph(paragraph._document, paragraph._paragraph, run) : null;
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(source ?? paragraph, paragraph,
                    tableDefaults, defaults, fontMap);
                double? size = source != null ? style.FontSize ?? (useConfiguredTypography ? null : defaults.FontSize) :
                    GetNativeParagraphStyleDefaults(paragraph).FontSize ?? tableDefaults.FontSize ?? (useConfiguredTypography ? null : defaults.FontSize);
                yield return style.ApplyTo(new PdfCore.PdfTextRun(numbersById.GetLabel(number),
                    bold: style.Bold, underline: style.Underline, italic: style.Italic, strike: style.Strike,
                    color: style.Color, fontSize: ((double?)(size) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null), font: style.Font, fontFamily: style.FontFamily,
                    baseline: PdfCore.PdfTextBaseline.Superscript, backgroundColor: style.BackgroundColor,
                    underlineStyle: style.UnderlineStyle, strikeStyle: style.StrikeStyle));
            }
        }
    }
}
