using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineTextDecorations(ICollection<HtmlRenderVisual> visuals,
        IDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals, HtmlInlineRun run,
        IElement? formattingContainer, IReadOnlyList<HtmlRenderVisual> textVisuals, bool aboveText) {
        HtmlRenderBoxStyle style = run.Style;
        if (!style.VectorTextDecoration) return;
        double thickness = style.DecorationThickness ?? Math.Max(1D, style.Font.Size / 16D);
        foreach (HtmlRenderText text in textVisuals.OfType<HtmlRenderText>()) {
            double width = Math.Max(0D, text.TextAdvanceWidth ?? text.TextPaintWidth ?? text.Width);
            if (width <= 0D) continue;
            double baseline = text.Y + style.Font.Size;
            if (aboveText) {
                Paint(style.StrikethroughStyle, baseline - style.Font.Size * 0.3D - thickness / 2D, "line-through");
            } else {
                Paint(style.UnderlineStyle, baseline + (style.UnderlineOffset ?? style.Font.Size / 10D), "underline");
                Paint(style.OverlineStyle, baseline - style.Font.Size - thickness / 2D, "overline");
            }

            void Paint(OfficeTextDecorationStyle pattern, double top, string line) {
                if (pattern == OfficeTextDecorationStyle.None) return;
                ChargeLayoutOperations(OfficeTextDecorationGeometry.GetCommandCount(width, thickness, pattern), run.Source ?? "text decoration");
                OfficeShape shape = OfficeTextDecorationGeometry.CreateHorizontalBand(width, thickness, pattern, style.DecorationColor);
                var visual = new HtmlRenderShape(shape, text.X, top, visuals.Count,
                    source: run.Source + ":decoration:" + line, layoutY: text.LayoutY, layoutHeight: text.LayoutHeight);
                AddInlineOwnedVisual(visuals, ownedVisuals,
                    visual.TranslateRelativePaint(run.PaintOffsetX, run.PaintOffsetY, visuals.Count), run.OwnerElement, formattingContainer);
            }
        }
    }
}
