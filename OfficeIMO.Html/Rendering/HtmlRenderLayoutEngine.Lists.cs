using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderFlowBlock ApplyListSemantics(HtmlRenderFlowBlock block, IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "ul" || tag == "ol") {
            return block.WithVisuals(new[] {
                new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.List,
                    0D,
                    0D,
                    Math.Max(0.01D, block.Width),
                    Math.Max(0.01D, block.Height),
                    block.Visuals,
                    0,
                    HtmlRenderStyleResolver.DescribeSource(element))
            });
        }

        if (tag != "li") return block;

        PartitionListMarkerVisuals(block.Visuals, out List<HtmlRenderVisual> markerVisuals, out List<HtmlRenderVisual> bodyVisuals);
        var itemVisuals = new List<HtmlRenderVisual>(2);
        if (markerVisuals.Count > 0) {
            (double x, double y, double width, double height) = ResolveSemanticBounds(markerVisuals, block.Width, block.Height);
            itemVisuals.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListLabel,
                x,
                y,
                width,
                height,
                markerVisuals,
                itemVisuals.Count,
                "list-marker"));
        }

        if (bodyVisuals.Count > 0) {
            itemVisuals.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListBody,
                0D,
                0D,
                Math.Max(0.01D, block.Width),
                Math.Max(0.01D, block.Height),
                bodyVisuals,
                itemVisuals.Count,
                HtmlRenderStyleResolver.DescribeSource(element)));
        }

        if (itemVisuals.Count == 0) return block;
        return block.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListItem,
                0D,
                0D,
                Math.Max(0.01D, block.Width),
                Math.Max(0.01D, block.Height),
                itemVisuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        });
    }

    private static void PartitionListMarkerVisuals(
        IEnumerable<HtmlRenderVisual> visuals,
        out List<HtmlRenderVisual> markerVisuals,
        out List<HtmlRenderVisual> bodyVisuals) {
        markerVisuals = new List<HtmlRenderVisual>();
        bodyVisuals = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            PartitionListMarkerVisual(visual, out HtmlRenderVisual? marker, out HtmlRenderVisual? body);
            if (marker != null) markerVisuals.Add(marker);
            if (body != null) bodyVisuals.Add(body);
        }
    }

    private static void PartitionListMarkerVisual(
        HtmlRenderVisual visual,
        out HtmlRenderVisual? marker,
        out HtmlRenderVisual? body) {
        if (string.Equals(visual.Source, "list-marker", StringComparison.Ordinal)) {
            marker = visual;
            body = null;
            return;
        }

        IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
        if (children == null) {
            marker = null;
            body = visual;
            return;
        }

        PartitionListMarkerVisuals(children, out List<HtmlRenderVisual> markerChildren, out List<HtmlRenderVisual> bodyChildren);
        marker = markerChildren.Count == 0 ? null : CloneGroupWithChildren(visual, markerChildren);
        body = bodyChildren.Count == 0 ? null : CloneGroupWithChildren(visual, bodyChildren);
    }

    private static IReadOnlyList<HtmlRenderVisual>? GetGroupChildren(HtmlRenderVisual visual) =>
        visual is HtmlRenderClipGroup clip ? clip.Visuals
            : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
            : visual is HtmlRenderEffectGroup effect ? effect.Visuals
            : visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
            : null;

    private static HtmlRenderVisual CloneGroupWithChildren(HtmlRenderVisual visual, IReadOnlyList<HtmlRenderVisual> children) {
        if (visual is HtmlRenderClipGroup clip) {
            return new HtmlRenderClipGroup(
                clip.ClipX,
                clip.ClipY,
                clip.ClipWidth,
                clip.ClipHeight,
                clip.ClipHorizontal,
                clip.ClipVertical,
                children,
                clip.PaintOrder,
                clip.Source,
                clip.LayoutY);
        }

        if (visual is HtmlRenderPathClipGroup pathClip) {
            return new HtmlRenderPathClipGroup(
                pathClip.ClipX,
                pathClip.ClipY,
                pathClip.ClipPath,
                children,
                pathClip.PaintOrder,
                pathClip.Source,
                pathClip.LayoutY);
        }

        if (visual is HtmlRenderEffectGroup effect) {
            return new HtmlRenderEffectGroup(
                effect.X,
                effect.Y,
                effect.Width,
                effect.Height,
                effect.Transform,
                effect.Opacity,
                children,
                effect.PaintOrder,
                effect.Source,
                effect.LayoutY);
        }

        HtmlRenderSemanticGroup semantic = (HtmlRenderSemanticGroup)visual;
        return new HtmlRenderSemanticGroup(
            semantic.Role,
            semantic.X,
            semantic.Y,
            semantic.Width,
            semantic.Height,
            children,
            semantic.PaintOrder,
            semantic.Source,
            semantic.ColumnSpan,
            semantic.RowSpan,
            semantic.HeaderScope,
            semantic.LayoutY);
    }

    private static (double X, double Y, double Width, double Height) ResolveSemanticBounds(
        IReadOnlyList<HtmlRenderVisual> visuals,
        double fallbackWidth,
        double fallbackHeight) {
        if (visuals.Count == 0) return (0D, 0D, Math.Max(0.01D, fallbackWidth), Math.Max(0.01D, fallbackHeight));
        double left = visuals.Min(visual => visual.X);
        double top = visuals.Min(visual => visual.Y);
        double right = visuals.Max(visual => visual.X + visual.Width);
        double bottom = visuals.Max(visual => visual.Y + visual.Height);
        return (left, top, Math.Max(0.01D, right - left), Math.Max(0.01D, bottom - top));
    }
}
