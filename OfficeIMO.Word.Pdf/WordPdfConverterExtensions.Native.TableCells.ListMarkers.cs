using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static (IReadOnlyList<PdfCore.PdfTextRun> Runs, double AnchorShift) CreateNativeCellListMarkerRuns(
            string marker, WordParagraph paragraph, NativeTableStyleDefaults tableStyleDefaults,
            NativeDocumentDefaults nativeDefaults, double firstLineIndent, NativeFontMap? nativeFontMap) {
            NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph,
                tableRunStyleDefaults: tableStyleDefaults.RunStyle, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null) {
                return (new[] { CreateNativeListMarkerTextRun(marker, paragraph, textStyle, nativeFontMap) }, 0D);
            }

            PdfCore.PdfTextRun markerRun = CreateNativeListMarkerTextRun(marker, paragraph, textStyle, nativeFontMap, includeSuffix: false);
            double markerFontSize = markerRun.FontSize ?? textStyle.FontSize ?? nativeDefaults.FontSize;
            NativeTextSpacing markerSpacing = ResolveNativeListMarkerTextSpacing(info.Value, textStyle.ListMarkerTextSpacing);
            double markerWidth = nativeFontMap?.MeasureText(markerRun)
                ?? EstimateNativeListMarkerWidth(marker, markerFontSize, markerSpacing);
            double anchorShift = GetNativeMarkerAnchorShift(info.Value, markerWidth);
            double suffixWidth = info.Value.LevelSuffix == WordListLevelSuffix.Space
                ? ResolveNativeListSpaceSuffixWidth(paragraph, nativeDefaults, nativeFontMap, tableStyleDefaults.RunStyle)
                : 0D;
            double trailingOffset = info.Value.LevelSuffix switch {
                WordListLevelSuffix.Nothing => 0D,
                WordListLevelSuffix.Space => suffixWidth,
                _ => Math.Max(0D, -firstLineIndent + anchorShift - markerWidth)
            };

            var result = new List<PdfCore.PdfTextRun>(2) { markerRun };
            AddNativeCellListSpacer(result, trailingOffset);
            return (result, anchorShift);
        }

        private static void AddNativeCellListSpacer(List<PdfCore.PdfTextRun> runs, double width) {
            if (width > 0.01D) {
                runs.Add(PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineBox(width, 0.01D, borderWidth: 0D)));
            }
        }
    }
}
