using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Word emits the numbering space in Arial at the marker's size,
        // independently of the marker and paragraph-mark font families.
        private static double ResolveNativeListSpaceSuffixWidth(WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap,
            NativeTableRunStyleDefaults tableRunStyleDefaults = default) {
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph,
                tableRunStyleDefaults: tableRunStyleDefaults, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            double fontSize = info?.MarkerFontSize ?? style.FontSize ?? nativeDefaults.FontSize;
            NativeTextSpacing spacing = info.HasValue
                ? ResolveNativeListMarkerTextSpacing(info.Value, style.ListMarkerTextSpacing) : style.ListMarkerTextSpacing;
            PdfCore.PdfStandardFont font = TryResolveNativeMappedFont("Arial", nativeFontMap, out var mappedFont)
                ? mappedFont : PdfCore.PdfStandardFont.Helvetica;
            string? family = nativeFontMap?.TryGetNamedFontFamily("Arial", out string? namedFamily) == true
                ? namedFamily : null;
            PdfCore.PdfTextRun space = spacing.ApplyTo(new PdfCore.PdfTextRun(" ",
                fontSize: fontSize, font: font, fontFamily: family));
            return Math.Max(0D, nativeFontMap?.MeasureText(space)
                ?? EstimateNativeListMarkerWidth(" ", fontSize, spacing));
        }
    }
}
