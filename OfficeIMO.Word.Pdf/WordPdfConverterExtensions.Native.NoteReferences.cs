using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Generated reference text must retain its source formatting independently
        // of the minimum font used to lay out the surrounding mixed-size paragraph.
        private static IEnumerable<PdfCore.PdfTextRun> CreateNativeNoteReferenceRuns(
            WordParagraph paragraph, IReadOnlyList<int> numbers, Dictionary<long, int> numbersById,
            NativeDocumentDefaults defaults, NativeFontMap? fontMap,
            NativeTableRunStyleDefaults tableDefaults = default, bool useConfiguredTypography = false) {
            var sourceKeys = new HashSet<long>();
            foreach (int number in numbers) {
                bool Matches(long key) => numbersById.TryGetValue(key, out int value) && value == number && sourceKeys.Add(key);
                WordParagraph? source = null;
                foreach (W.Run run in paragraph._paragraph.Descendants<W.Run>()) {
                    bool matches = run.Elements<W.FootnoteReference>().Any(reference =>
                        reference.Id?.Value is long id && Matches(GetNativeFootnoteKey(id))) ||
                        run.Elements<W.EndnoteReference>().Any(reference =>
                        reference.Id?.Value is long id && Matches(GetNativeEndnoteKey(id)));
                    if (matches) {
                        source = new WordParagraph(paragraph._document, paragraph._paragraph, run);
                        break;
                    }
                }
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(source ?? paragraph, paragraph,
                    tableDefaults, defaults, fontMap);
                double? size = source != null ? style.FontSize ?? (useConfiguredTypography ? null : defaults.FontSize) :
                    GetNativeParagraphStyleDefaults(paragraph).FontSize ?? tableDefaults.FontSize ?? (useConfiguredTypography ? null : defaults.FontSize);
                yield return new PdfCore.PdfTextRun(number.ToString(CultureInfo.InvariantCulture),
                    bold: style.Bold, underline: style.Underline, italic: style.Italic, strike: style.Strike,
                    color: style.Color, fontSize: size, font: style.Font, fontFamily: style.FontFamily,
                    baseline: PdfCore.PdfTextBaseline.Superscript, backgroundColor: style.BackgroundColor,
                    underlineStyle: style.UnderlineStyle, strikeStyle: style.StrikeStyle);
            }
        }
    }
}
