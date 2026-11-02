using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Trailing spaces and tabs do not produce a metric-bearing wrapped
        // segment. Authored line breaks do, and must retain their own lines.
        private static bool IsNativeTrailingSpacingOnly(string text) =>
            text.All(character => character == ' ' || character == '\t');

        private static bool EndsWithNativeLineBreak(string text) =>
            text.TrimEnd(' ', '\t').EndsWith("\n", StringComparison.Ordinal);

        private static bool NeedsNativeParagraphMarkStrut(System.Collections.Generic.IEnumerable<PdfCore.PdfTextRun> runs, bool numbered) {
            PdfCore.PdfTextRun? finalContent = runs.LastOrDefault(run =>
                run.InlineElement != null || !IsNativeTrailingSpacingOnly(run.Text));
            return finalContent == null ? numbered :
                finalContent.InlineElement == null && EndsWithNativeLineBreak(finalContent.Text);
        }

        private static PdfCore.PdfTextRun CreateNativeParagraphMarkStrut(WordParagraph paragraph,
            NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            // Resolve a detached view of the mark through the ordinary typography
            // owner. The first text run may be empty, hidden or differently styled;
            // it cannot supply the paragraph mark's font, size or emphasis.
            var properties = new W.RunProperties();
            W.ParagraphMarkRunProperties? mark = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
            if (mark != null) {
                foreach (var property in mark.ChildElements) properties.Append(property.CloneNode(true));
            }
            var markView = new WordParagraph(paragraph._document, paragraph._paragraph, new W.Run(properties));
            PdfCore.PdfTextRun strut = CreateNativeCellTextRun(string.Empty, markView,
                tableStyleDefaults, nativeDefaults, nativeFontMap);
            strut.IsLineBoxStrut = true;
            return strut;
        }
    }
}
