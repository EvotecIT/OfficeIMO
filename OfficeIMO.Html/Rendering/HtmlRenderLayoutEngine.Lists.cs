using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderFlowBlock ApplyListSemantics(HtmlRenderFlowBlock block, IElement element, string? structureElementKey = null, int? logicalOrder = null) {
        string tag = element.TagName.ToLowerInvariant();
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        if (tag == "ul" || tag == "ol") {
            return WrapSemanticBlock(block, HtmlRenderSemanticGroupRole.List, source,
                structureElementKey == null ? null : structureElementKey + ":list", logicalOrder);
        }

        if (tag != "li") return block;
        return WrapSemanticFragments(block, (visuals, height) =>
            CreateListItemSemanticVisuals(visuals, block.Width, height, source, structureElementKey, logicalOrder));
    }

    private static IReadOnlyList<HtmlRenderVisual> CreateListItemSemanticVisuals(
        IReadOnlyList<HtmlRenderVisual> visuals, double width, double height, string source, string? structureElementKey, int? logicalOrder) {
        PartitionListMarkerVisuals(visuals, out List<HtmlRenderVisual> markerVisuals, out List<HtmlRenderVisual> bodyVisuals);
        var itemVisuals = new List<HtmlRenderVisual>(2);
        if (markerVisuals.Count > 0) {
            if (markerVisuals.Count == 1
                && markerVisuals[0] is HtmlRenderSemanticGroup existingLabel
                && existingLabel.Role == HtmlRenderSemanticGroupRole.ListLabel) {
                markerVisuals = existingLabel.Visuals.ToList();
            }
            (double x, double y, double markerWidth, double markerHeight) = ResolveSemanticBounds(markerVisuals, width, height);
            itemVisuals.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListLabel, x, y, markerWidth, markerHeight,
                markerVisuals, itemVisuals.Count, "list-marker",
                structureElementKey: structureElementKey == null ? null : structureElementKey + ":label"));
        }
        if (bodyVisuals.Count > 0) {
            itemVisuals.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListBody, 0D, 0D, Math.Max(0.01D, width), Math.Max(0.01D, height),
                bodyVisuals, itemVisuals.Count, source,
                structureElementKey: structureElementKey == null ? null : structureElementKey + ":body"));
        }
        if (itemVisuals.Count == 0) return visuals;
        return new[] {
            new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.ListItem, 0D, 0D, Math.Max(0.01D, width), Math.Max(0.01D, height),
                itemVisuals, 0, source,
                structureElementKey: structureElementKey == null ? null : structureElementKey + ":item", logicalOrder: logicalOrder)
        };
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
            : visual is HtmlRenderLayoutRegion region ? region.Visuals
            : visual is HtmlRenderLogicalTextGroup logicalText ? logicalText.Visuals
            : null;

    private static HtmlRenderVisual CloneGroupWithChildren(HtmlRenderVisual visual, IReadOnlyList<HtmlRenderVisual> children) =>
        visual.CopyStackingContextTo(CloneGroupWithChildrenCore(visual, children));

    private static HtmlRenderVisual CloneGroupWithChildrenCore(HtmlRenderVisual visual, IReadOnlyList<HtmlRenderVisual> children) {
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
                clip.LayoutY, clip.IsViewportOverflow, clip.IsFlowFragment, clip.LegacyClipOwnerNodeId);
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

        if (visual is HtmlRenderSemanticGroup semantic) {
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
                semantic.LayoutY,
                semantic.StructureElementKey, semantic.LayoutHeight, semantic.AlternativeText, semantic.MathMlSource, semantic.LogicalOrder);
        }


        if (visual is HtmlRenderLayoutRegion region) {
            var rebuiltRegion = new HtmlRenderLayoutRegion(
                region.SourceKey,
                region.RegionKind,
                region.SourceText,
                region.Position,
                region.FloatSide,
                region.ZIndex,
                region.BackgroundLayerCount,
                region.BoxShadowLayerCount,
                region.BackgroundColor,
                region.X,
                region.Y,
                region.Width,
                region.Height,
                children,
                region.PaintOrder,
                region.Source,
                region.LayoutY);
            rebuiltRegion.SurfaceNumber = region.SurfaceNumber;
            rebuiltRegion.SemanticSectionNumber = region.SemanticSectionNumber;
            rebuiltRegion.SemanticSectionOriginX = region.SemanticSectionOriginX;
            rebuiltRegion.SemanticSectionOriginY = region.SemanticSectionOriginY;
            rebuiltRegion.SemanticTableNumber = region.SemanticTableNumber;
            rebuiltRegion.SemanticTableOriginX = region.SemanticTableOriginX;
            rebuiltRegion.SemanticTableOriginY = region.SemanticTableOriginY;
            return rebuiltRegion;
        }

        HtmlRenderLogicalTextGroup logicalText = (HtmlRenderLogicalTextGroup)visual;
        return new HtmlRenderLogicalTextGroup(
            logicalText.Text,
            logicalText.X,
            logicalText.Y,
            logicalText.Width,
            logicalText.Height,
            children,
            logicalText.PaintOrder,
            logicalText.Source,
            logicalText.LayoutY,
            logicalText.LayoutHeight,
            logicalText.LogicalScope,
            logicalText.IsFlowAnchor);
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
