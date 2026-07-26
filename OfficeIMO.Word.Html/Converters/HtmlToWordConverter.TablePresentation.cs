using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void ApplyTableBackground(
            WordTableCell cell,
            string value,
            string? ancestorBackdrop) {
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
                string.IsNullOrEmpty(cell.ShadingFillColorHex)
                    ? ancestorBackdrop
                    : cell.ShadingFillColorHex);
        }

        private static string? ResolveAncestorBlockBackground(AngleSharp.Dom.IElement element) {
            var lineage = new Stack<AngleSharp.Dom.IElement>();
            for (AngleSharp.Dom.IElement? current = element.ParentElement;
                 current != null;
                 current = current.ParentElement) {
                lineage.Push(current);
            }

            string? backdrop = null;
            while (lineage.Count > 0) {
                AngleSharp.Dom.IElement ancestor = lineage.Pop();
                string? resolved = ResolveBlockBackground(
                    ancestor.GetAttribute("style") ?? string.Empty,
                    backdrop);
                if (!string.IsNullOrEmpty(resolved)) {
                    backdrop = resolved;
                }
            }
            return backdrop;
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
