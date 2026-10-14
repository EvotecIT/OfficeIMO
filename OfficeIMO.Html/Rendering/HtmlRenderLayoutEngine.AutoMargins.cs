using AngleSharp.Dom;

namespace OfficeIMO.Html {

    internal sealed partial class HtmlRenderLayoutEngine {
        private HtmlRenderBoxStyle ResolveNormalFlowHorizontalAutoMargins(IElement element, HtmlRenderBoxStyle style, double containingWidth) {
            if (!style.MarginLeftAuto && !style.MarginRightAuto) return style;
            if (element.LocalName.Equals("table", StringComparison.OrdinalIgnoreCase) && !style.ExplicitWidth.HasValue) {
                // An auto-width table needs its intrinsic columns before its free space is known.
                return style;
            }

            double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
            double boxWidth = IsReplacedImageElement(element)
                ? ResolveReplacedImageBoxWidth(element, style)
                : ResolveBoxWidth(availableWidth, style);
            double freeSpace = Math.Max(0D, containingWidth - style.MarginLeft - style.MarginRight - boxWidth);
            if (freeSpace <= 0D) return style;

            var resolved = style.Clone();
            ApplyHorizontalAutoMarginFreeSpace(resolved, freeSpace);
            return resolved;
        }

        /// <summary>Records the computed share separately from authored margins and float-band placement.</summary>
        private static void ApplyHorizontalAutoMarginFreeSpace(HtmlRenderBoxStyle style, double freeSpace) {
            double left = style.MarginLeftAuto ? (style.MarginRightAuto ? freeSpace / 2D : freeSpace) : 0D;
            double right = style.MarginRightAuto ? (style.MarginLeftAuto ? freeSpace / 2D : freeSpace) : 0D;
            style.MarginLeft += left;
            style.MarginRight += right;
            style.HorizontalAutoMarginLeftShare += left;
            style.HorizontalAutoMarginRightShare += right;
        }

        /// <summary>Restores the placement baseline before sizing a table again, including a reused style.</summary>
        private static void RemoveHorizontalAutoMarginShares(HtmlRenderBoxStyle style) {
            style.MarginLeft -= style.HorizontalAutoMarginLeftShare;
            style.MarginRight -= style.HorizontalAutoMarginRightShare;
            style.HorizontalAutoMarginLeftShare = 0D;
            style.HorizontalAutoMarginRightShare = 0D;
        }
    }
}
