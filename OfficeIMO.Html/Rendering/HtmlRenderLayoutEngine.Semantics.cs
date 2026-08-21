using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock ApplyElementSemantics(HtmlRenderFlowBlock block, IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
        int nodeId = GetSemanticNodeId(element);
        string structureElementKey = "html-element:" + nodeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (ShouldAssignNavigationNode(style) && !style.BookmarkSuppressed) {
            string anchorText = ResolveBookmarkAnchorText(element, style);
            if (anchorText.Length > 0 && !ContainsBookmarkAnchor(block.Visuals, nodeId)) {
                block = block.WithVisuals(block.Visuals.Concat(new HtmlRenderVisual[] {
                    new HtmlRenderBookmarkAnchor(
                        nodeId,
                        anchorText,
                        style.MarginLeft,
                        style.MarginTop,
                        0.01D,
                        0.01D,
                        block.Visuals.Count,
                        HtmlRenderStyleResolver.DescribeSource(element))
                }));
            }
        }
        HtmlRenderFlowBlock listBlock = style.SemanticArtifact || style.SemanticGroupRoleOverride.HasValue
            ? block
            : ApplyListSemantics(block, element, structureElementKey);
        HtmlRenderSemanticGroupRole role;
        if (style.SemanticArtifact) role = HtmlRenderSemanticGroupRole.Artifact;
        else if (style.SemanticGroupRoleOverride.HasValue) role = style.SemanticGroupRoleOverride.Value;
        else if (!TryResolveSemanticGroupRole(element.TagName, out role)) return WrapEditableLayoutRegion(listBlock, element, style);
        HtmlRenderFlowBlock semanticBlock = listBlock.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                role,
                0D,
                0D,
                Math.Max(0.01D, listBlock.Width),
                Math.Max(0.01D, listBlock.Height),
                listBlock.Visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element),
                structureElementKey: structureElementKey)
        });
        return WrapEditableLayoutRegion(semanticBlock, element, style);
    }

    private HtmlRenderFlowBlock ApplySpecializedElementSemantics(HtmlRenderFlowBlock block, IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
        int nodeId = GetSemanticNodeId(element);
        string structureElementKey = "html-element:" + nodeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (ShouldAssignNavigationNode(style) && !style.BookmarkSuppressed) {
            string anchorText = ResolveBookmarkAnchorText(element, style);
            if (anchorText.Length > 0) {
                block = block.WithVisuals(block.Visuals.Concat(new HtmlRenderVisual[] {
                    new HtmlRenderBookmarkAnchor(
                        nodeId,
                        anchorText,
                        0D,
                        0D,
                        0.01D,
                        0.01D,
                        block.Visuals.Count,
                        HtmlRenderStyleResolver.DescribeSource(element))
                }));
            }
        }
        if (!style.SemanticArtifact && !style.SemanticGroupRoleOverride.HasValue) return WrapEditableLayoutRegion(block, element, style);
        HtmlRenderSemanticGroupRole role = style.SemanticArtifact
            ? HtmlRenderSemanticGroupRole.Artifact
            : style.SemanticGroupRoleOverride!.Value;
        HtmlRenderFlowBlock semanticBlock = block.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                role,
                0D,
                0D,
                Math.Max(0.01D, block.Width),
                Math.Max(0.01D, block.Height),
                block.Visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element),
                structureElementKey: structureElementKey)
        });
        return WrapEditableLayoutRegion(semanticBlock, element, style);
    }

    private HtmlRenderFlowBlock WrapEditableLayoutRegion(
        HtmlRenderFlowBlock block,
        IElement element,
        HtmlRenderBoxStyle style,
        HtmlRenderLayoutRegionKind? forcedKind = null) {
        if (!_options.EnableEditableLayoutRegions) return block;
        string? sourceKey = _suppressedEditableLayoutRegionMarkers.Contains(element)
            ? null
            : HtmlEditableLayoutProjector.GetRegionSourceKey(element);
        if (string.IsNullOrWhiteSpace(sourceKey)) return block;
        if (!style.PaintVisible) return block;
        if (!forcedKind.HasValue
            && style.Position != "absolute" && style.Position != "fixed"
            && style.FloatSide != "left" && style.FloatSide != "right"
            && style.Display != "flex" && style.Display != "inline-flex"
            && style.Display != "grid" && style.Display != "inline-grid") return block;
        HtmlRenderLayoutRegionKind kind = forcedKind ?? ResolveEditableLayoutRegionKind(style);
        int zIndex = int.TryParse(style.ZIndex, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out int parsedZIndex) ? parsedZIndex : 0;
        double availableBoxWidth = Math.Max(1D, block.Width - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableBoxWidth, style);
        double boxHeight = style.ExplicitHeight.HasValue || style.AspectRatio.HasValue
            ? ResolveBoxHeight(Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom), boxWidth, style)
            : Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom);
        return block.WithVisuals(new[] {
            new HtmlRenderLayoutRegion(
                sourceKey!,
                kind,
                CollapseFlexText(ResolveLogicalText(block.Visuals, string.Empty)),
                style.Position,
                style.FloatSide,
                zIndex,
                style.BackgroundImageLayerCount,
                style.BoxShadowLayerCount,
                style.BackgroundColor,
                style.MarginLeft,
                style.MarginTop,
                boxWidth,
                boxHeight,
                block.Visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        });
    }

    private static HtmlRenderLayoutRegionKind ResolveEditableLayoutRegionKind(HtmlRenderBoxStyle style) {
        if (style.Position == "absolute" || style.Position == "fixed") return HtmlRenderLayoutRegionKind.Positioned;
        if (style.FloatSide == "left" || style.FloatSide == "right") return HtmlRenderLayoutRegionKind.Floating;
        return style.Display == "grid" || style.Display == "inline-grid"
            ? HtmlRenderLayoutRegionKind.Grid
            : HtmlRenderLayoutRegionKind.Flex;
    }

    private HtmlRenderFlowBlock LayoutElementWithoutEditableRegionMarker(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth) {
        string? sourceKey = HtmlEditableLayoutProjector.GetRegionSourceKey(element);
        if (sourceKey == null) return LayoutElement(element, containingWidth, style, parentStyle, depth);
        _suppressedEditableLayoutRegionMarkers.Add(element);
        try {
            return LayoutElement(element, containingWidth, style, parentStyle, depth);
        } finally {
            _suppressedEditableLayoutRegionMarkers.Remove(element);
        }
    }

    private IReadOnlyList<HtmlRenderFlowBlock> ApplyFlattenedElementSemantics(
        IReadOnlyList<HtmlRenderFlowBlock> blocks,
        FlattenedSemanticBoundary boundary) {
        if (blocks.Count == 0) return blocks;
        var result = new List<HtmlRenderFlowBlock>(blocks.Count);
        for (int index = 0; index < blocks.Count; index++) {
            result.Add(ApplyFlattenedSemanticBoundary(blocks[index], boundary, index == 0));
        }
        return result;
    }

    private FlattenedSemanticBoundary CreateFlattenedSemanticBoundary(IElement element, HtmlRenderBoxStyle style) {
        if (_flattenedSemanticBoundaries.TryGetValue(element, out FlattenedSemanticBoundary? existing)) return existing;
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
        int nodeId = GetSemanticNodeId(element);
        bool hasNavigation = ShouldAssignNavigationNode(style) && !style.BookmarkSuppressed;
        HtmlRenderSemanticGroupRole? role = style.SemanticArtifact
            ? HtmlRenderSemanticGroupRole.Artifact
            : style.SemanticGroupRoleOverride;
        if (!role.HasValue && TryResolveSemanticGroupRole(element.TagName, out HtmlRenderSemanticGroupRole naturalRole)) {
            role = naturalRole;
        }
        var boundary = new FlattenedSemanticBoundary(
            element,
            style,
            role,
            nodeId,
            hasNavigation ? ResolveBookmarkAnchorText(element, style) : string.Empty,
            "html-display-contents:" + nodeId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _flattenedSemanticBoundaries[element] = boundary;
        return boundary;
    }

    private HtmlRenderFlowBlock ApplyFlattenedSemanticBoundary(
        HtmlRenderFlowBlock block,
        FlattenedSemanticBoundary boundary,
        bool firstFragment) {
        string anchorText = boundary.AnchorText.Length > 0
            ? boundary.AnchorText
            : ResolveVisibleBookmarkText(boundary.Element);
        if (firstFragment
            && anchorText.Length > 0
            && !ContainsBookmarkAnchor(block.Visuals, boundary.NodeId)) {
            block = block.WithVisuals(block.Visuals.Concat(new HtmlRenderVisual[] {
                new HtmlRenderBookmarkAnchor(
                    boundary.NodeId,
                    anchorText,
                    0D,
                    0D,
                    0.01D,
                    0.01D,
                    block.Visuals.Count,
                    boundary.Source)
            }));
        }

        HtmlRenderFlowBlock semanticBlock = boundary.Style.SemanticArtifact || boundary.Style.SemanticGroupRoleOverride.HasValue
            ? block
            : ApplyListSemantics(block, boundary.Element, boundary.StructureElementKey);
        if (!boundary.Role.HasValue) return semanticBlock;
        return semanticBlock.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                boundary.Role.Value,
                0D,
                0D,
                Math.Max(0.01D, semanticBlock.Width),
                Math.Max(0.01D, semanticBlock.Height),
                semanticBlock.Visuals,
                0,
                boundary.Source,
                structureElementKey: boundary.StructureElementKey)
        });
    }

    private sealed class FlattenedSemanticBoundary {
        internal FlattenedSemanticBoundary(
            IElement element,
            HtmlRenderBoxStyle style,
            HtmlRenderSemanticGroupRole? role,
            int nodeId,
            string anchorText,
            string structureElementKey) {
            Element = element;
            Style = style;
            Role = role;
            NodeId = nodeId;
            AnchorText = anchorText;
            StructureElementKey = structureElementKey;
            Source = HtmlRenderStyleResolver.DescribeSource(element);
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal HtmlRenderSemanticGroupRole? Role { get; }
        internal int NodeId { get; }
        internal string AnchorText { get; }
        internal string StructureElementKey { get; }
        internal string Source { get; }
    }

    private sealed class FlattenedSemanticPlacement {
        internal FlattenedSemanticPlacement(FlattenedSemanticBoundary boundary, bool firstFragment) {
            Boundary = boundary;
            FirstFragment = firstFragment;
        }

        internal FlattenedSemanticBoundary Boundary { get; }
        internal bool FirstFragment { get; }
    }

    private string ResolveBookmarkAnchorText(IElement element, HtmlRenderBoxStyle style) {
        if (!string.IsNullOrWhiteSpace(style.BookmarkLabel)) return style.BookmarkLabel!.Trim();
        string renderedText = ResolveVisibleBookmarkText(element, out bool rootVisible);
        if (renderedText.Length > 0) return renderedText;
        if (!rootVisible) return string.Empty;
        return CollapseFlexText(element.GetAttribute("aria-label")
            ?? element.GetAttribute("alt")
            ?? element.GetAttribute("title")
            ?? string.Empty);
    }

    private string ResolveVisibleBookmarkText(IElement element) => ResolveVisibleBookmarkText(element, out _);

    private string ResolveVisibleBookmarkText(IElement element, out bool rootVisible) {
        if (ShouldSkipElement(element)
            || !TryResolveBookmarkTextState(element, inheritedVisibility: true, out bool visible, out bool prunesSubtree)
            || prunesSubtree) {
            rootVisible = false;
            return string.Empty;
        }
        rootVisible = visible;
        return CollapseFlexText(string.Concat(EnumerateVisibleBookmarkText(element.ChildNodes, visible)));
    }

    private IEnumerable<string> EnumerateVisibleBookmarkText(IEnumerable<INode> nodes, bool inheritedVisibility) {
        foreach (INode node in nodes) {
            if (node is IText text) {
                if (inheritedVisibility) yield return text.Data;
                continue;
            }
            if (node is not IElement element
                || ShouldSkipElement(element)
                || !TryResolveBookmarkTextState(element, inheritedVisibility, out bool visible, out bool prunesSubtree)
                || prunesSubtree) {
                continue;
            }
            foreach (string childText in EnumerateVisibleBookmarkText(element.ChildNodes, visible)) yield return childText;
        }
    }

    private bool TryResolveBookmarkTextState(IElement element, bool inheritedVisibility, out bool visible, out bool prunesSubtree) {
        if (_layoutStyles.TryGetValue(element, out HtmlRenderBoxStyle? layoutStyle)) {
            visible = layoutStyle.PaintVisible;
            prunesSubtree = layoutStyle.Display == "none" || layoutStyle.SemanticArtifact;
            return true;
        }
        if (!_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computedStyle)) {
            visible = inheritedVisibility;
            prunesSubtree = false;
            return true;
        }

        string visibility = computedStyle.GetValue("visibility").Trim().ToLowerInvariant();
        visible = visibility == "visible"
            ? true
            : visibility == "hidden" || visibility == "collapse"
                ? false
                : inheritedVisibility;
        prunesSubtree = string.Equals(computedStyle.GetValue("display"), "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(computedStyle.GetValue("-officeimo-pdf-tag-type"), "artifact", StringComparison.OrdinalIgnoreCase)
            || string.Equals(computedStyle.GetValue("-officeimo-pdf-tag-type"), "none", StringComparison.OrdinalIgnoreCase);
        return true;
    }

    private HtmlRenderVisual ApplyInlineElementSemantics(HtmlRenderVisual visual, HtmlInlineRun run) {
        if (!run.InlineSemanticGroupRole.HasValue) return visual;
        HtmlRenderSemanticGroupRole role = run.InlineSemanticGroupRole.Value;
        if (visual is HtmlRenderSemanticGroup existing && existing.Role == role) return visual;
        return new HtmlRenderSemanticGroup(
            role,
            visual.X,
            visual.Y,
            Math.Max(0.01D, visual.Width),
            Math.Max(0.01D, visual.Height),
            new[] { visual },
            visual.PaintOrder,
            run.Source,
            layoutY: visual.LayoutY,
            structureElementKey: run.InlineSemanticGroupKey);
    }

    private static bool ContainsBookmarkAnchor(IEnumerable<HtmlRenderVisual> visuals, int semanticNodeId) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderBookmarkAnchor anchor && anchor.SemanticNodeId == semanticNodeId) return true;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals
                : null;
            if (children != null && ContainsBookmarkAnchor(children, semanticNodeId)) return true;
        }
        return false;
    }

    private void RegisterInlineSemanticControls(IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
    }

    private void ReportUnsupportedSemanticTag(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedSemanticTag.Length == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.PdfSemanticTagUnsupported,
            "A PDF semantic tag override was invalid and automatic HTML semantics were used.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            "-officeimo-pdf-tag-type=" + style.UnsupportedSemanticTag,
            OfficeConversionLossKind.Approximation);
    }

    private void RegisterBookmark(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedBookmark.Length > 0) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.BookmarkValueUnsupported,
                "A CSS bookmark value used automatic heading navigation.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(element),
                style.UnsupportedBookmark,
                OfficeConversionLossKind.Approximation);
        }
        if (style.SemanticArtifact) return;
        bool automaticHeading = HtmlRenderHeading.TryGetLevel(style.SemanticRole, out int headingLevel);
        if (!automaticHeading && !style.BookmarkLevel.HasValue) return;
        int nodeId = GetSemanticNodeId(element);
        int level = style.BookmarkLevel ?? headingLevel;
        _bookmarkDefinitions[nodeId] = new HtmlRenderBookmarkDefinition(level, style.BookmarkLabel, style.BookmarkState, style.BookmarkSuppressed, GetDocumentOrder(element));
    }

    private static bool TryResolveSemanticGroupRole(string tagName, out HtmlRenderSemanticGroupRole role) {
        string tag = tagName.ToLowerInvariant();
        if (tag == "p") {
            role = HtmlRenderSemanticGroupRole.Paragraph;
            return true;
        }
        if (tag == "h1") {
            role = HtmlRenderSemanticGroupRole.Heading1;
            return true;
        }
        if (tag == "h2") {
            role = HtmlRenderSemanticGroupRole.Heading2;
            return true;
        }
        if (tag == "h3") {
            role = HtmlRenderSemanticGroupRole.Heading3;
            return true;
        }
        if (tag == "h4") {
            role = HtmlRenderSemanticGroupRole.Heading4;
            return true;
        }
        if (tag == "h5") {
            role = HtmlRenderSemanticGroupRole.Heading5;
            return true;
        }
        if (tag == "h6") {
            role = HtmlRenderSemanticGroupRole.Heading6;
            return true;
        }
        if (tag == "main" || tag == "section" || tag == "article" || tag == "nav" || tag == "aside") {
            role = HtmlRenderSemanticGroupRole.Section;
            return true;
        }
        if (tag == "header" || tag == "footer") {
            role = HtmlRenderSemanticGroupRole.Division;
            return true;
        }

        role = default;
        return false;
    }
}