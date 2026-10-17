using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Paints the collapsed gap between matching paragraph fills inside the first panel without changing text placement.</summary>
        private static PdfCore.PdfParagraphStyle JoinNativeAdjacentParagraphShading(WordParagraph? nextParagraph,
            PdfCore.PdfParagraphStyle paragraphStyle, PdfCore.PdfPanelStyle panelStyle, NativeDocumentDefaults nativeDefaults) {
            if (nextParagraph == null || nextParagraph.IsPageBreak || HasNativePageBreakBefore(nextParagraph) ||
                string.IsNullOrEmpty(nextParagraph.Text) || GetHeadingLevel(nextParagraph) > 0 ||
                !panelStyle.Background.HasValue || HasNativePanelBorders(panelStyle)) return paragraphStyle;
            PdfCore.PdfParagraphStyle nextStyle = CreateNativeParagraphStyle(nextParagraph, nativeDefaults);
            PdfCore.PdfPanelStyle? nextPanel = CreateNativeParagraphPanelStyle(nextParagraph, nextStyle);
            if (nextPanel == null || !Equals(nextPanel.Background, panelStyle.Background) || HasNativePanelBorders(nextPanel) ||
                nextStyle.LeftIndent != paragraphStyle.LeftIndent || nextStyle.RightIndent != paragraphStyle.RightIndent) return paragraphStyle;
            PdfCore.PdfParagraphStyle result = paragraphStyle.Clone();
            result.SpacingAfter = Math.Max(panelStyle.SpacingAfter, nextStyle.SpacingBefore);
            panelStyle.SpacingAfter = 0D;
            return result;
        }

        private static bool HasNativePanelBorders(PdfCore.PdfPanelStyle style) => style.BorderColor.HasValue ||
            style.TopBorder != null || style.RightBorder != null || style.BottomBorder != null || style.LeftBorder != null;
    }
}
