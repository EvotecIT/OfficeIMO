using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // A numbering space follows the paragraph mark's typography, not the
        // numbering symbol's font or the first visible body's direct formatting.
        private static WordParagraph CreateNativeParagraphMarkSource(WordParagraph paragraph) {
            var properties = new W.RunProperties();
            W.ParagraphMarkRunProperties? mark = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
            if (mark != null) {
                foreach (var child in mark.ChildElements) properties.Append(child.CloneNode(true));
            }
            // The temporary run is never attached to the source document.
            return new WordParagraph(paragraph._document, paragraph._paragraph, new W.Run(properties));
        }

        private static double ResolveNativeListSpaceSuffixWidth(WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap,
            NativeTableRunStyleDefaults tableRunStyleDefaults = default) {
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(CreateNativeParagraphMarkSource(paragraph),
                tableRunStyleDefaults: tableRunStyleDefaults, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            PdfCore.PdfTextRun space = style.TextSpacing.ApplyTo(new PdfCore.PdfTextRun(" ",
                bold: style.Bold, italic: style.Italic, fontSize: style.FontSize,
                font: style.Font, fontFamily: style.FontFamily));
            return Math.Max(0D, nativeFontMap?.MeasureText(space)
                ?? EstimateNativeListMarkerWidth(" ", style.FontSize ?? nativeDefaults.FontSize, style.TextSpacing));
        }
    }
}
