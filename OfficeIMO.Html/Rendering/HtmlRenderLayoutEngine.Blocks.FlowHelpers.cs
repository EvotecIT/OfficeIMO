using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderBoxStyle SuppressContinuationStartDecorations(HtmlRenderBoxStyle style) {
        HtmlRenderBoxStyle continuation = style.Clone();
        continuation.MarginTop = 0D;
        continuation.PaddingTop = 0D;
        continuation.Borders = new HtmlRenderBorderEdges(
            continuation.Borders.Top.WithWidth(0D),
            continuation.Borders.Right,
            continuation.Borders.Bottom,
            continuation.Borders.Left);
        continuation.BreakBefore = HtmlPageBreakTarget.None;
        continuation.StringSet = string.Empty;
        return continuation;
    }

    private HtmlRenderFlowBlock AttachElementMargins(HtmlRenderFlowBlock block, HtmlRenderBoxStyle style, IElement element,
        bool collapsesThrough = false, HtmlCollapsedMargin? topGroup = null, HtmlCollapsedMargin? bottomGroup = null) {
        HtmlRenderFlowBlock attached = block.WithCollapsibleMargins(style.MarginTop, style.MarginBottom, element,
            collapsesThrough, topGroup, bottomGroup);
        attached.AvoidBreakInsideScope = style.AvoidBreakInsideScope != HtmlFragmentationScope.None
            ? style.AvoidBreakInsideScope
            : style.AvoidBreakInside ? HtmlFragmentationScope.All : HtmlFragmentationScope.None;
        IReadOnlyList<HtmlCssRunningStringAssignment> ownAssignments =
            ResolveRunningStringAssignments(element, style, 0D);
        return ownAssignments.Count == 0
            ? attached
            : attached.WithRunningStringAssignments(ownAssignments.Concat(attached.RunningStringAssignments));
    }

    private IReadOnlyList<HtmlCssRunningStringAssignment> ResolveRunningStringAssignments(
        IElement element,
        HtmlRenderBoxStyle style,
        double offset) {
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        IReadOnlyList<HtmlCssRunningStringAssignment> ownAssignments = HtmlCssRunningStringParser.ResolveAssignments(
            element,
            style.StringSet,
            _options.MaxRunningStringCharacters,
            count => ChargeLayoutOperations(count, source),
            (contentElement, maximumCharacters, chargeOperations) =>
                ResolveRunningStringContentText(
                    contentElement,
                    style,
                    maximumCharacters,
                    chargeOperations),
            out bool limitExceeded);
        if (limitExceeded) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.RunningStringLimitExceeded,
                "A CSS running-string value exceeded the configured character limit and was omitted.",
                HtmlDiagnosticSeverity.Warning,
                source,
                "limit=" + _options.MaxRunningStringCharacters);
        }
        int documentOrder = GetDocumentOrder(element);
        ownAssignments = ownAssignments
            .Select(assignment => assignment.InDocumentOrder(documentOrder))
            .ToList()
            .AsReadOnly();
        return Math.Abs(offset) <= 0.0001D
            ? ownAssignments
            : ownAssignments.Select(assignment => assignment.Translate(offset)).ToList().AsReadOnly();
    }

    private static bool CanCollapseParentMargin(HtmlRenderBoxStyle style, bool top) {
        if (style.Display != "block" && style.Display != "list-item") return false;
        if (style.OverflowX != "visible" || style.OverflowY != "visible") return false;
        if (top) return style.BorderTopWidth <= 0D && style.PaddingTop <= 0D;
        return style.BorderBottomWidth <= 0D
            && style.PaddingBottom <= 0D
            && !style.ExplicitHeight.HasValue
            && (!style.MinHeight.HasValue || style.MinHeight.Value <= 0D);
    }

    private static bool CanCollapseThroughEmptyBlock(
        HtmlRenderBoxStyle style,
        bool usesBlockFormatting,
        IReadOnlyList<HtmlRenderFlowBlock> children,
        IReadOnlyList<HtmlRenderVisual> contentVisuals,
        double contentHeight,
        bool hasInFlowLineBoxes) {
        if (style.Display != "block" || style.OverflowX != "visible" || style.OverflowY != "visible") return false;
        if (style.BorderTopWidth > 0D || style.BorderBottomWidth > 0D || style.PaddingTop > 0D || style.PaddingBottom > 0D) return false;
        if (style.ExplicitHeight.HasValue || style.AspectRatio.HasValue || style.MinHeight.HasValue && style.MinHeight.Value > 0D) return false;
        if (usesBlockFormatting) return children.Count == 0 || children.All(child => child.CollapsesThrough);
        return !hasInFlowLineBoxes && contentVisuals.Count == 0 && contentHeight <= 0.0001D;
    }

    private static bool CanUseZeroHeightForMarginCollapse(HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, double contentHeight) =>
        contentHeight <= 0.0001D
        && style.Display == "block"
        && parentStyle.Display != "flex"
        && parentStyle.Display != "inline-flex"
        && parentStyle.Display != "grid"
        && parentStyle.Display != "inline-grid"
        && style.BorderTopWidth <= 0D
        && style.BorderBottomWidth <= 0D
        && style.PaddingTop <= 0D
        && style.PaddingBottom <= 0D
        && !style.ExplicitHeight.HasValue
        && !style.AspectRatio.HasValue
        && (!style.MinHeight.HasValue || style.MinHeight.Value <= 0D);

    private double FlushInlineNodes(ICollection<HtmlRenderFlowBlock> blocks, List<INode> nodes, double width, HtmlRenderBoxStyle style, IElement sourceElement, int depth, out bool hasInFlowLineBoxes, PagedFloatBoundary? pageBoundary = null, ICollection<HtmlFloatExclusion>? activeFloats = null, ICollection<HtmlFloatExclusion>? emittedFloats = null, double flowOffset = 0D, InlineFloatContext? floatContext = null) {
        hasInFlowLineBoxes = false;
        if (nodes.Count == 0) return 0D;
        HtmlInlineLayout inline = LayoutInlineNodes(nodes, width, style, depth + 1, null, null,
            pageBoundary: pageBoundary,
            inheritedFloats: activeFloats?.Count > 0 ? activeFloats.Select(item => item.Shift(0D, -flowOffset)).ToArray() : null,
            applyTextIndent: blocks.Count == 0, floatContext: floatContext);
        if (activeFloats != null) {
            foreach (HtmlFloatExclusion exclusion in inline.FloatExclusions) {
                HtmlFloatExclusion positioned = exclusion.Shift(0D, flowOffset);
                activeFloats.Add(positioned);
                emittedFloats?.Add(positioned);
            }
        }
        if (inline.Height > inline.NormalFlowHeight + 0.0001D) {
            _inlineFloatOverhangs.TryGetValue(sourceElement, out double previousOverhang);
            _inlineFloatOverhangs[sourceElement] = Math.Max(
                previousOverhang,
                inline.Height - inline.NormalFlowHeight);
        }
        hasInFlowLineBoxes = inline.HasInFlowLineBoxes == true
            || inline.InterruptedFlow?.HasZeroHeightOrdinaryContent == true;
        nodes.Clear();
        if (inline.InterruptedFlow != null) {
            blocks.Add(inline.InterruptedFlow);
            return inline.Height;
        }
        if (inline.Visuals.Count == 0 && !hasInFlowLineBoxes) return 0D;
        // A floated sibling does not consume normal-flow height before the
        // following block. A block formatting context still contains its float.
        bool containsFloat = EstablishesFloatContainingBlock(style);
        double flowHeight = floatContext != null ? inline.NormalFlowHeight : activeFloats != null && (!containsFloat || HasMultiColumnLayout(style))
            ? inline.NormalFlowHeight : inline.Height;
        var block = new HtmlRenderFlowBlock(
            width,
            flowHeight,
            inline.Visuals,
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            false,
            HtmlRenderStyleResolver.DescribeSource(sourceElement),
            inline.BreakOffsets,
            inline.LineBreakOffsets,
            style.Orphans,
            style.Widows,
            lineBreakGroups: inline.LineBreakGroups,
            continuationGroups: inline.ContinuationGroups,
            trailingGroups: inline.TrailingGroups,
            pageName: style.PageName,
            runningStringAssignments: inline.RunningStringAssignments,
            forcedBreaks: inline.ForcedBreaks,
            layoutViewportWidth: ActiveSurfaceWidth,
            layoutViewportHeight: _activePageGeometry.Height,
            pagedPaintExtent: inline.PagedPaintExtent,
            avoidBreakRanges: inline.AvoidBreakRanges,
            safeBreakRanges: inline.SafeBreakRanges);
        block.HasZeroHeightOrdinaryContent = flowHeight == 0D && inline.HasInFlowLineBoxes == true;
        blocks.Add(block);
        return block.Height;
    }

    private static bool EstablishesFloatContainingBlock(HtmlRenderBoxStyle style) =>
        style.Display is "flow-root" or "flex" or "inline-flex" or "grid" or "inline-grid"
        || style.ColumnCount.HasValue || style.ColumnWidth.HasValue
        || style.OverflowX is "auto" or "hidden" or "scroll"
        || style.OverflowY is "auto" or "hidden" or "scroll";

    private bool HasBlockChildren(IElement element, double width, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, element);
        if (IncludeColumnGeneratedContent(element, HtmlPseudoElementKind.Before)
                && HasBlockGeneratedContent(element, HtmlPseudoElementKind.Before, width, parentStyle)
            || IncludeColumnGeneratedContent(element, HtmlPseudoElementKind.After)
                && HasBlockGeneratedContent(element, HtmlPseudoElementKind.After, width, parentStyle)) return true;
        foreach (IElement child in SelectedColumnChildNodes(element).OfType<IElement>()) {
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(child, width, parentStyle);
            if (style.FloatSide != "none") return true;
            if (style.Display == "contents" && HasBlockChildren(child, width, style, depth + 1)) return true;
            if (style.Display != "none" && ShouldExtractOutOfFlow(style) && !UsesInlineStaticPosition(child, style)) return true;
            if (style.Display != "none" && HtmlRenderStyleResolver.IsBlockElement(child, style)) return true;
            if (style.Display != "none" && ContainsFloatingDescendant(child, width, style, depth + 1)) return true;
        }

        return false;
    }

    private bool HasBlockGeneratedContent(IElement element, HtmlPseudoElementKind kind, double width, HtmlRenderBoxStyle parentStyle) {
        if (!_generatedContent.TryGetContent(element, kind, out _)
            || !_styleResolver.TryResolvePseudo(element, kind, width, parentStyle, out HtmlRenderBoxStyle style)) return false;
        if (CanPositionGeneratedContentLocally(style, parentStyle)) return false;
        return style.Display is "block" or "flow-root" or "list-item" or "flex" or "grid";
    }

    private static bool UsesInlineStaticPosition(IElement element, HtmlRenderBoxStyle style) {
        if (style.Display == "inline-flex" || style.Display == "inline-grid" || style.Display == "inline-block") return true;
        if (style.Display != "inline") return false;
        return style.DisplayWasSpecified || !HtmlRenderStyleResolver.IsDefaultBlockElement(element);
    }

    private HtmlRenderFlowBlock LayoutHorizontalRule(IElement element, double containingWidth, HtmlRenderBoxStyle style) {
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double width = ResolveBoxWidth(availableWidth, style);
        double lineWidth = style.BorderWidth > 0D ? style.BorderWidth : 1D;
        var shape = OfficeShape.Rectangle(width, lineWidth);
        shape.FillColor = style.BorderColor;
        shape.StrokeWidth = 0D;
        double height = style.MarginTop + lineWidth + style.MarginBottom;
        var visual = new HtmlRenderShape(shape, style.MarginLeft, style.MarginTop, 0, source: HtmlRenderStyleResolver.DescribeSource(element));
        IReadOnlyList<HtmlRenderVisual> visuals = style.PaintVisible ? new[] { visual } : Array.Empty<HtmlRenderVisual>();
        return new HtmlRenderFlowBlock(containingWidth, Math.Max(height, 0.01D), visuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, HtmlRenderStyleResolver.DescribeSource(element), pageName: style.PageName);
    }

    private double ResolveBoxWidth(double availableWidth, HtmlRenderBoxStyle style) {
        double width = style.ExplicitWidth ?? (style.BorderBox ? availableWidth : Math.Max(1D, availableWidth - style.HorizontalInsets));
        if (!style.BorderBox) width += style.HorizontalInsets;
        if (style.MaxWidth.HasValue) width = Math.Min(width, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MinWidth.HasValue) width = Math.Max(width, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        return Math.Max(style.HorizontalInsets, Math.Max(1D, width));
    }

    private static double ResolveBoxHeight(double contentHeight, double boxWidth, HtmlRenderBoxStyle style, bool allowZeroHeight = false) {
        double height;
        bool usesAspectRatio = !style.ExplicitHeight.HasValue && style.AspectRatio.HasValue;
        if (style.ExplicitHeight.HasValue) {
            height = style.ExplicitHeight.Value;
        } else if (style.AspectRatio.HasValue) {
            double ratioWidth = style.BorderBox
                ? boxWidth
                : Math.Max(0D, boxWidth - style.HorizontalInsets);
            height = ratioWidth / style.AspectRatio.Value;
        } else {
            height = contentHeight;
        }
        if (!style.BorderBox || !style.ExplicitHeight.HasValue && !usesAspectRatio) height += style.VerticalInsets;
        if (style.MaxHeight.HasValue) height = Math.Min(height, style.MaxHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
        if (style.MinHeight.HasValue) height = Math.Max(height, style.MinHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
        return Math.Max(style.VerticalInsets, Math.Max(allowZeroHeight ? 0D : 0.01D, height));
    }

    private void ReportUnsupportedLayout(IElement element, HtmlRenderBoxStyle style) {
        string display = style.Display;
        if (display == "flex" || display == "inline-flex") {
            AddUnsupported(HtmlRenderDiagnosticCodes.FlexLayoutPending, "This flex formatting case is not active yet; children use normal flow.", element);
        } else if (display == "grid" || display == "inline-grid") {
            AddUnsupported(HtmlRenderDiagnosticCodes.GridLayoutPending, "Grid layout is not yet active in the direct HTML renderer; children use normal flow.", element);
        }
    }

    private bool ShouldSkipElement(IElement element) {
        if (IsClosedDisclosureChild(element)) return true;
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "input" && string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase)) return true;
        return tag == "head" || tag == "style" || tag == "script" || tag == "template" || tag == "noscript" || tag == "meta" || tag == "link" || tag == "title" || tag == "base";
    }

    private void EnsureDepth(int depth, IElement element) {
        if (depth <= _options.MaxLayoutDepth) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.DepthLimitExceeded,
            "HTML layout exceeded the configured maximum depth at " + HtmlRenderStyleResolver.DescribeSource(element) + ".",
            nameof(HtmlRenderOptions.MaxLayoutDepth),
            depth,
            _options.MaxLayoutDepth);
    }

    private HtmlListMarker? ResolveListMarker(IElement element, HtmlRenderBoxStyle style, double containingWidth) {
        bool hasPseudoStyle = _styleResolver.TryResolvePseudo(
            element, HtmlPseudoElementKind.Marker, containingWidth, style, out HtmlRenderBoxStyle markerStyle);
        if (_generatedContent.Suppresses(element, HtmlPseudoElementKind.Marker)) return null;

        string? markerContent = null;
        if (_generatedContent.TryGetContent(element, HtmlPseudoElementKind.Marker, out HtmlGeneratedContent generatedContent)
            && generatedContent.Fragments.Any(fragment => fragment.Kind != HtmlGeneratedContentFragmentKind.Text)
            && TryCreateGeneratedListMarker(element, markerStyle, containingWidth, generatedContent, out HtmlRenderFlowBlock generatedMarkerBlock)) {
            return new HtmlListMarker(string.Empty, markerStyle, style.ListStylePosition, generatedMarkerBlock);
        }
        if (_generatedContent.TryGet(element, HtmlPseudoElementKind.Marker, out string generatedMarker)) {
            markerContent = ApplyTextTransform(generatedMarker, markerStyle);
        }
        if (markerContent == null && TryCreateListImageMarker(element, style, hasPseudoStyle ? markerStyle : style, containingWidth, out HtmlRenderFlowBlock imageMarker)) {
            return new HtmlListMarker(string.Empty, hasPseudoStyle ? markerStyle : style, style.ListStylePosition, imageMarker);
        }
        if (markerContent == null && string.Equals(style.ListStyleType, "none", StringComparison.OrdinalIgnoreCase)) return null;
        IElement? parent = element.ParentElement;
        if (markerContent == null && parent == null) markerContent = "• ";
        bool ordered = parent != null && string.Equals(parent.TagName, "ol", StringComparison.OrdinalIgnoreCase);
        string listStyle = style.ListStyleType.Length == 0 ? ordered ? "decimal" : "disc" : style.ListStyleType;
        int ordinal = HtmlListSemantics.TryResolveOrdinal(element, out int resolvedOrdinal) ? resolvedOrdinal : 1;
        if (markerContent == null && _counterStyles.TryFormatMarker(ordinal, listStyle, out string customMarker, out bool customMarkerLimited)) {
            if (customMarkerLimited) ReportCounterRepresentationLimit(element, listStyle);
            markerContent = customMarker.Length == 0 ? null : customMarker;
        }
        if (markerContent == null) {
            if (!HtmlCounterStyleFormatter.TryFormat(ordinal, listStyle, out string marker, out bool markerLimited)) {
                marker = ordered ? ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) : "•";
            }
            if (markerLimited) ReportCounterRepresentationLimit(element, listStyle);
            if (marker.Length > 0) markerContent = marker + HtmlCounterStyleFormatter.MarkerSuffix(markerLimited ? "decimal" : listStyle);
        }
        if (string.IsNullOrEmpty(markerContent)) return null;
        return new HtmlListMarker(markerContent!, hasPseudoStyle ? markerStyle : style, style.ListStylePosition);
    }

    private bool TryCreateGeneratedListMarker(
        IElement element,
        HtmlRenderBoxStyle markerStyle,
        double containingWidth,
        HtmlGeneratedContent content,
        out HtmlRenderFlowBlock marker) {
        string source = DescribePseudoSource(element, HtmlPseudoElementKind.Marker);
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineFragments(content, element, markerStyle, null, source, containingWidth, 0D, 0D, runs);
        runs = ApplyScopedFontFallbacks(runs);
        if (runs.Count == 0) {
            marker = null!;
            return false;
        }

        HtmlInlineLayout inline = LayoutInlineRuns(runs, Math.Max(1D, containingWidth * 0.5D), markerStyle, element);
        (double left, double top, double width, double height) = ResolveSemanticBounds(
            inline.Visuals,
            markerStyle.LineHeight,
            inline.Height);
        var label = new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.ListLabel,
            left,
            top,
            width,
            height,
            inline.Visuals,
            0,
            "list-marker");
        marker = new HtmlRenderFlowBlock(
            width,
            Math.Max(0.01D, inline.Height),
            new HtmlRenderVisual[] { label },
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            source);
        return true;
    }

    private bool TryCreateListImageMarker(
        IElement element,
        HtmlRenderBoxStyle listStyle,
        HtmlRenderBoxStyle markerStyle,
        double containingWidth,
        out HtmlRenderFlowBlock marker) {
        marker = null!;
        if (string.Equals(listStyle.ListStyleImage, "none", StringComparison.OrdinalIgnoreCase)) return false;
        IReadOnlyList<string> sources = HtmlResourcePipeline.ExtractCssUrls(listStyle.ListStyleImage);
        if (sources.Count != 1) return false;
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(element) + "::marker:list-style-image";
        if (!TryResolveImageSource(sources[0], sourceDescription, out byte[]? bytes, out _, out OfficeImageInfo? imageInfo)
            || bytes == null || imageInfo == null) return false;
        IDocument? owner = element.Owner;
        if (owner == null) return false;
        IElement imageElement = owner.CreateElement("img");
        imageElement.SetAttribute("src", sources[0]);
        var imageStyle = new HtmlRenderBoxStyle {
            Display = "inline",
            PaintVisible = markerStyle.PaintVisible,
            Font = markerStyle.Font,
            Color = markerStyle.Color,
            LineHeight = markerStyle.LineHeight,
            UsesNaturalLineHeight = markerStyle.UsesNaturalLineHeight,
            InlineAlignmentLineHeight = markerStyle.InlineAlignmentLineHeight,
            SemanticRole = "list-marker",
            ApplyEmbeddedImageOrientation = listStyle.ApplyEmbeddedImageOrientation,
            ImageResolutionDpi = listStyle.ImageResolutionDpi
        };
        double imageWidth = ResolveFloatingImageOuterWidth(imageElement, imageStyle);
        HtmlRenderFlowBlock image = LayoutImage(imageElement, imageWidth, imageStyle);
        var markerVisual = new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.ListLabel,
            0D,
            0D,
            Math.Max(0.01D, image.Width),
            Math.Max(0.01D, image.Height),
            image.Visuals,
            0,
            "list-marker");
        marker = new HtmlRenderFlowBlock(
            image.Width,
            image.Height,
            new[] { markerVisual },
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            sourceDescription);
        return true;
    }

    private void ReportCounterRepresentationLimit(IElement element, string style) {
        AddUnsupported(
            HtmlRenderDiagnosticCodes.CounterRepresentationLimitExceeded,
            "A CSS counter representation exceeded the managed rendering budget and used a decimal fallback.",
            element,
            "list-style-type=" + style,
            OfficeConversionLossKind.Approximation);
    }
}
