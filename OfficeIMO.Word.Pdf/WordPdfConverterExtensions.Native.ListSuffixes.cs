using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static double ResolveNativeListTabBodyPosition(WordParagraph paragraph,
            double markerEnd, double textIndent, NativeDocumentDefaults nativeDefaults) {
            if (markerEnd <= textIndent + 0.01D) return textIndent;
            var stops = new Dictionary<int, WordTabStop>();
            W.Tabs? levelTabs = WordDocumentTraversal.GetListInfo(paragraph)?.LevelTabStops;
            if (levelTabs != null) {
                foreach (W.TabStop tab in levelTabs.Elements<W.TabStop>().Take(MaxNativeParagraphTabStops)) {
                    var stop = new WordTabStop(paragraph, (W.TabStop)tab.CloneNode(true));
                    stops[stop.Position] = stop;
                }
            }
            foreach (WordTabStop stop in GetNativeParagraphEffectiveTabStops(paragraph)) stops[stop.Position] = stop;
            foreach (WordTabStop tabStop in stops.Values.Where(tab => IsNativeRenderableTextTabStop(tab.Alignment))
                .OrderBy(tab => tab.Position)) {
                double position = tabStop.Position / 20D;
                if (position > markerEnd + 0.01D) return position;
            }
            double interval = nativeDefaults.DefaultTabStopWidth ?? 36D;
            return (Math.Floor(markerEnd / interval) + 1D) * interval;
        }

        private static double ResolveNativeInlineListMarkerColumnWidth(WordParagraph paragraph,
            string marker, PdfCore.PdfParagraphStyle paragraphStyle,
            NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            double width = Math.Max(0D, -paragraphStyle.FirstLineIndent);
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (marker.Length == 0 || info == null || info.Value.LevelSuffix is WordListLevelSuffix.Space or WordListLevelSuffix.Nothing)
                return width;
            NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph,
                nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            PdfCore.PdfTextRun markerRun = CreateNativeListMarkerTextRun(marker, paragraph, textStyle, nativeFontMap, includeSuffix: false);
            double markerWidth = nativeFontMap.MeasureText(markerRun)
                ?? EstimateNativeListMarkerWidth(marker, markerRun.FontSize ?? nativeDefaults.FontSize,
                    ResolveNativeListMarkerTextSpacing(info.Value, textStyle.ListMarkerTextSpacing));
            double markerStart = paragraphStyle.LeftIndent + paragraphStyle.FirstLineIndent;
            double bodyPosition = ResolveNativeListTabBodyPosition(paragraph, markerStart + markerWidth,
                paragraphStyle.LeftIndent, nativeDefaults);
            return Math.Max(width, bodyPosition - markerStart);
        }

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
