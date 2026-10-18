using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Paint-only page window, including repeated groups. A completed table's backdrop
    /// stops after its final footer; a longer ancestor can still continue to the page edge.
    /// </summary>
    private readonly record struct PagedBackgroundWindow(
        double End, double? CompletedSourceEnd, double CompletedPaintEnd) {
        internal double ResolveEnd(double sourceBottom) => CompletedSourceEnd.HasValue
            && sourceBottom <= CompletedSourceEnd.Value + 0.0001D
                ? Math.Min(End, CompletedPaintEnd) : End;
    }

    private static PagedBackgroundWindow ResolvePagedBackgroundWindow(double sourceStart, PagedBlockFragment plan) =>
        new(sourceStart + plan.Available + plan.FooterHeight,
            plan.Footer != null && plan.End >= plan.Footer.ContentEndsAt - 0.0001D
                ? plan.Footer.SourceEndsAt : null,
            plan.End + plan.FooterHeight);

    /// <summary>Retains the table's own plain backdrop beneath a detached repeated header.</summary>
    private static void PrependRepeatedTableBackdrop(
        List<HtmlRenderVisual> visuals, HtmlRenderShape? background, double height) {
        if (background == null || visuals.Count == 0 || height <= 0.0001D) return;
        OfficeShape fill = OfficeShape.Rectangle(background.Width, height);
        fill.FillColor = background.InnerShape.FillColor;
        fill.StrokeWidth = 0D;
        HtmlRenderShape repeated = background.CopyStackingContextTo(new HtmlRenderShape(fill,
            background.X, 0D, 0, source: background.Source,
            fragmentSideBorders: background.FragmentSideBorders));
        var ordered = new List<HtmlRenderVisual> { repeated };
        AppendFragmentSideBorders(ordered, repeated, 0D, 0D, height);
        foreach (HtmlRenderVisual visual in visuals) {
            ordered.Add(visual is HtmlRenderSemanticGroup group
                ? visual.CopyStackingContextTo(group.ProjectPaint(group.Visuals, 0D, 0D, ordered.Count))
                : visual.Translate(0D, 0D, ordered.Count));
        }
        visuals.Clear();
        visuals.AddRange(ordered);
    }

    private static bool CanExtendBoxBackgroundThroughFragmentSlack(HtmlRenderBoxStyle style, OfficeShape fill) =>
        (style.Display is "block" or "flex" or "grid" or "flow-root" or "table")
        // Flex/grid stretch records a used height; the authored auto-height box still fragments.
        && (!style.ExplicitHeight.HasValue || style.AutoHeightLayoutStretch) && !style.AspectRatio.HasValue
        && style.Position == "static" && style.FloatSide == "none"
        && style.Opacity >= 1D && fill.FillColor?.A == byte.MaxValue
        && fill.Kind == OfficeShapeKind.Rectangle && (!style.HasBorderLayout || style.Display == "table"
            || CanExtendOrdinarySliceSides(style))
        && style.BoxShadowLayerCount == 0 && !style.HasDeclaredBackgroundImage
        && style.BackgroundColorClip == "border-box" && style.ClipPath == "none"
        && style.UsedTransform == "none" && style.UsedIndividualScale == "none";

    /// <summary>
    /// Continues only a box's own solid side bands beside its background. Closing
    /// edges and collapsed grid rules retain their source cut and normal flow.
    /// </summary>
    private static void AppendFragmentSideBorders(
        ICollection<HtmlRenderVisual> fragment,
        HtmlRenderShape background,
        double fragmentStart,
        double contentEnd,
        double backgroundEnd) {
        if (background.FragmentSideBorders is not HtmlRenderBorderEdges borders) return;
        AddSide(borders.Left, left: true);
        AddSide(borders.Right, left: false);

        void AddSide(HtmlRenderBorderSide side, bool left) {
            if (!side.IsPainted || side.Style != "solid") return;
            double width = Math.Min(background.Width, side.Width);
            OfficeShape band = OfficeShape.Rectangle(width, backgroundEnd - contentEnd);
            band.FillColor = side.Color;
            band.StrokeWidth = 0D;
            HtmlRenderShape visual = background.CopyStackingContextTo(new HtmlRenderShape(band,
                left ? background.X : background.X + background.Width - width,
                contentEnd, background.PaintOrder,
                source: background.Source + (left ? ":fragment-border-left" : ":fragment-border-right"),
                layoutY: background.LayoutY, layoutHeight: background.LayoutHeight));
            fragment.Add(CreateVerticallyClippedVisualFragment(visual, fragmentStart,
                contentEnd, backgroundEnd, fragment.Count));
        }
    }

    private static bool CanExtendOrdinarySliceSides(HtmlRenderBoxStyle style) =>
        style.Display != "table" && style.BoxDecorationBreak == "slice"
        && (!style.Borders.Left.IsPainted || style.Borders.Left.Style == "solid")
        && (!style.Borders.Right.IsPainted || style.Borders.Right.Style == "solid");

    private void ExtendPagedBodyBackgroundThroughFragmentSlack(
        HtmlRenderFlowBlock block,
        HtmlCssPageGeometry geometry,
        ICollection<HtmlRenderVisual> visuals,
        double fragmentBottom) {
        IElement? body = _document.Body;
        if (body == null || !ReferenceEquals(block.OwnerElement, body) || ReferenceEquals(_surfaceRootElement, body)) return;
        if (!_layoutStyles.TryGetValue(body, out HtmlRenderBoxStyle? style)
            || !style.PaintVisible
            || !style.BackgroundColor.HasValue
            || style.BackgroundColor.Value.A != byte.MaxValue
            || style.HasDeclaredBackgroundImage
            || style.Opacity < 1D
            || style.HorizontalInsets > 0.0001D
            || style.VerticalInsets > 0.0001D
            || style.HasBorderLayout
            || !string.Equals(style.BackgroundColorClip, "border-box", StringComparison.Ordinal)
            || !string.Equals(style.BorderRadius, "0", StringComparison.Ordinal)
            || style.BorderTopLeftRadius.Length > 0 || style.BorderTopRightRadius.Length > 0
            || style.BorderBottomLeftRadius.Length > 0 || style.BorderBottomRightRadius.Length > 0
            || !string.Equals(style.UsedTransform, "none", StringComparison.Ordinal)
            || !string.Equals(style.UsedIndividualScale, "none", StringComparison.Ordinal)
            || !string.Equals(style.ClipPath, "none", StringComparison.Ordinal)
            || Math.Abs(style.MarginLeft) > 0.0001D || Math.Abs(style.MarginRight) > 0.0001D) return;
        if (geometry.PrintProduction != null
            || geometry.Margins.Left > 0.0001D || geometry.Margins.Right > 0.0001D
            || geometry.Margins.Top > 0.0001D || geometry.Margins.Bottom > 0.0001D
            || Math.Abs(block.Width - geometry.ContentWidth) > 0.0001D) return;

        double slack = geometry.Height - fragmentBottom;
        if (slack <= 0.0001D) return;
        // A body fragment continues on the next page even when a safe break leaves unused
        // space here. Its background covers that slack, but not the last page after body ends.
        OfficeShape fill = OfficeShape.Rectangle(geometry.Width, slack);
        fill.FillColor = style.BackgroundColor.Value;
        fill.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(fill, 0D, fragmentBottom, int.MinValue + 1024,
            source: "body:fragment-background-continuation"));
    }
}
