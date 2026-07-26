using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void ApplyTableBackground(WordTableCell cell, string value) {
            CssStyleMapper.CssProperties parsed =
                CssStyleMapper.ParseStyles("background-color:" + value);
            if (string.IsNullOrEmpty(parsed.BackgroundColor)) {
                return;
            }

            double alpha = parsed.BackgroundColorAlpha ?? 1d;
            if (alpha <= 0d) {
                return;
            }

            cell.ShadingFillColorHex = ResolveOpaqueTextBackground(
                parsed.BackgroundColor!,
                alpha,
                cell.ShadingFillColorHex);
        }

        private static void ClearSyntheticTableTrailingSpacing(Paragraph paragraph) {
            if (!HasOnlySyntheticZeroSpacing(paragraph)) {
                return;
            }

            SpacingBetweenLines? spacing =
                paragraph.ParagraphProperties?.GetFirstChild<SpacingBetweenLines>();
            spacing?.Remove();
        }
    }
}
