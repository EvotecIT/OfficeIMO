using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool IsRootLayoutContainer(IElement element) => ReferenceEquals(element, _document.Body ?? _document.DocumentElement);

    private static bool ShouldExtractOutOfFlow(HtmlRenderBoxStyle childStyle) {
        return childStyle.Position == "absolute" || childStyle.Position == "fixed";
    }

    private void RegisterOutOfFlowElement(
        IElement container,
        IElement element,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth) {
        if (style.Position == "fixed") {
            if (_registeredFixedElements.Add(element)) _fixedPositionedElements.Add(new PositionedElementRequest(element, style.Clone(), parentStyle.Clone(), depth));
            return;
        }
        if (style.Position != "absolute" || !_registeredAbsoluteElements.Add(element)) return;

        var request = new PositionedElementRequest(element, style.Clone(), parentStyle.Clone(), depth);
        IElement containingBlock = ResolveAbsoluteContainingBlock(element, container, parentStyle);
        if (IsRootLayoutContainer(containingBlock)) {
            _rootPositionedElements.Add(request);
            return;
        }

        if (!_localPositionedElements.TryGetValue(containingBlock, out List<PositionedElementRequest>? requests)) {
            requests = new List<PositionedElementRequest>();
            _localPositionedElements[containingBlock] = requests;
        }
        requests.Add(request);
    }

    private IElement ResolveAbsoluteContainingBlock(IElement element, IElement directParent, HtmlRenderBoxStyle directParentStyle) {
        IElement root = _document.Body ?? _document.DocumentElement ?? directParent;
        if (ReferenceEquals(directParent, root)) return root;
        if (EstablishesSupportedAbsoluteContainingBlock(directParent, directParentStyle)) return directParent;
        if (directParentStyle.Position != "static") return ReportInlineContainingBlockFallback(element, directParent, root);

        double referenceWidth = Math.Max(1D, (_options.Mode == HtmlRenderMode.Paged ? _options.PageWidth : _options.ViewportWidth) - _options.Margins.Left - _options.Margins.Right);
        bool flattenedFlexOrGridChild = directParentStyle.Display == "contents";
        for (IElement? ancestor = directParent.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (ReferenceEquals(ancestor, root)) return root;
            HtmlRenderBoxStyle ancestorStyle = _styleResolver.Resolve(ancestor, referenceWidth);
            if (ancestorStyle.Position != "static") {
                return EstablishesSupportedAbsoluteContainingBlock(ancestor, ancestorStyle)
                    ? ancestor
                    : ReportInlineContainingBlockFallback(element, ancestor, root);
            }
            if (flattenedFlexOrGridChild && (ancestorStyle.Display == "flex" || ancestorStyle.Display == "grid" || ancestorStyle.Display == "inline-flex" || ancestorStyle.Display == "inline-grid")) {
                return ancestor;
            }
            flattenedFlexOrGridChild = flattenedFlexOrGridChild && ancestorStyle.Display == "contents";
        }

        return root;
    }

    private static bool EstablishesSupportedAbsoluteContainingBlock(IElement element, HtmlRenderBoxStyle style) {
        if (style.Display == "flex" || style.Display == "grid" || style.Display == "inline-flex" || style.Display == "inline-grid") return true;
        return style.Position != "static" && HtmlRenderStyleResolver.IsBlockElement(element, style);
    }

    private IElement ReportInlineContainingBlockFallback(IElement element, IElement containingBlock, IElement root) {
        if (_reportedPositionContainingBlockFallbacks.Add(element)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
                "An inline positioned containing block is not active; the absolute element used the initial containing block.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(element),
                "containing-block=" + HtmlRenderStyleResolver.DescribeSource(containingBlock));
        }
        return root;
    }

    private void AppendLocalPositionedVisuals(
        IElement container,
        double containingWidth,
        double containingHeight,
        double originX,
        double originY,
        ICollection<HtmlRenderVisual> visuals) {
        if (!_localPositionedElements.TryGetValue(container, out List<PositionedElementRequest>? requests)) return;
        foreach (PositionedElementRequest request in requests) {
            PositionedLayer layer = request.Resolve(this, containingWidth, containingHeight);
            foreach (HtmlRenderVisual visual in layer.Block.Visuals) {
                visuals.Add(visual.Translate(originX + layer.X, originY + layer.Y, visuals.Count));
            }
        }
    }

    private void AppendPositionedRequests(
        ICollection<HtmlRenderVisual> visuals,
        IEnumerable<PositionedElementRequest> requests,
        double containingWidth,
        double containingHeight,
        double originX,
        double originY) {
        if (requests is IList<PositionedElementRequest> requestList) {
            for (int index = 0; index < requestList.Count; index++) {
                AppendPositionedRequest(visuals, requestList[index], containingWidth, containingHeight, originX, originY);
            }
            return;
        }

        foreach (PositionedElementRequest request in requests) {
            AppendPositionedRequest(visuals, request, containingWidth, containingHeight, originX, originY);
        }
    }

    private void AppendPositionedRequest(
        ICollection<HtmlRenderVisual> visuals,
        PositionedElementRequest request,
        double containingWidth,
        double containingHeight,
        double originX,
        double originY) {
        PositionedLayer layer = request.Resolve(this, containingWidth, containingHeight);
        foreach (HtmlRenderVisual visual in layer.Block.Visuals) {
            visuals.Add(visual.Translate(originX + layer.X, originY + layer.Y, _paintOrder++));
        }
    }

    private PositionedLayer LayoutPositionedElement(
        IElement element,
        HtmlRenderBoxStyle sourceStyle,
        HtmlRenderBoxStyle parentStyle,
        double containingWidth,
        double containingHeight,
        int depth) {
        HtmlRenderBoxStyle style = sourceStyle.Clone();
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        double? left = ResolveOutOfFlowInset(style.Left, containingWidth, style, source, "left");
        double? right = ResolveOutOfFlowInset(style.Right, containingWidth, style, source, "right");
        double? top = ResolveOutOfFlowInset(style.Top, containingHeight, style, source, "top");
        double? bottom = ResolveOutOfFlowInset(style.Bottom, containingHeight, style, source, "bottom");
        double outerWidth = ResolvePositionedOuterWidth(element, style, containingWidth, left, right);
        if (!style.ExplicitWidth.HasValue) SetPositionedExplicitWidth(style, outerWidth);
        if (!style.ExplicitHeight.HasValue && top.HasValue && bottom.HasValue) {
            double targetOuterHeight = Math.Max(0.01D, containingHeight - top.Value - bottom.Value);
            double targetBoxHeight = Math.Max(0.01D, targetOuterHeight - style.MarginTop - style.MarginBottom);
            style.ExplicitHeight = style.BorderBox ? targetBoxHeight : Math.Max(0.01D, targetBoxHeight - style.VerticalInsets);
        }
        if (style.ZIndex != "auto") {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.PositionZIndexPending, "CSS z-index is not yet active; the positioned element paints after in-flow content.", HtmlDiagnosticSeverity.Warning, source, "z-index=" + style.ZIndex);
        }
        style.Position = "static";
        style.ZIndex = "auto";
        HtmlRenderFlowBlock block = LayoutElement(element, Math.Max(1D, outerWidth), style, parentStyle, depth);
        double x = left ?? (right.HasValue ? containingWidth - right.Value - block.Width : 0D);
        double y = top ?? (bottom.HasValue ? containingHeight - bottom.Value - block.Height : 0D);
        return new PositionedLayer(block, x, y);
    }

    private double ResolvePositionedOuterWidth(IElement element, HtmlRenderBoxStyle style, double containingWidth, double? left, double? right) {
        if (style.ExplicitWidth.HasValue) {
            double available = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
            return Math.Min(containingWidth, style.MarginLeft + ResolveBoxWidth(available, style) + style.MarginRight);
        }
        if (left.HasValue && right.HasValue) return Math.Max(1D, containingWidth - left.Value - right.Value);
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "table") return containingWidth;
        double contentWidth = tag == "img"
            ? 300D
            : Math.Max(1D, MeasureText(ApplyTextTransform(CollapseFlexText(element.TextContent), style.TextTransform), style.Font));
        return Math.Max(1D, Math.Min(containingWidth, contentWidth + style.HorizontalInsets + style.MarginLeft + style.MarginRight));
    }

    private static void SetPositionedExplicitWidth(HtmlRenderBoxStyle style, double targetOuterWidth) {
        double targetBoxWidth = Math.Max(0.01D, targetOuterWidth - style.MarginLeft - style.MarginRight);
        style.ExplicitWidth = style.BorderBox ? targetBoxWidth : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
    }

    private double? ResolveOutOfFlowInset(string value, double reference, HtmlRenderBoxStyle style, string source, string property) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) return null;
        if (HtmlRenderCssValues.TryLength(value, reference, style.Font.Size, _options.DefaultFontSize, out double resolved)) return resolved;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.PositionInsetUnsupported, "A positioned inset could not be resolved and used auto.", HtmlDiagnosticSeverity.Warning, source, property + "=" + value);
        return null;
    }

    private sealed class PositionedElementRequest {
        private PositionedLayer? _cached;
        private double _width;
        private double _height;
        internal PositionedElementRequest(IElement element, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth) {
            Element = element;
            Style = style;
            ParentStyle = parentStyle;
            Depth = depth;
        }
        private IElement Element { get; }
        private HtmlRenderBoxStyle Style { get; }
        private HtmlRenderBoxStyle ParentStyle { get; }
        private int Depth { get; }
        internal PositionedLayer Resolve(HtmlRenderLayoutEngine engine, double width, double height) {
            if (_cached == null || Math.Abs(width - _width) > 0.0001D || Math.Abs(height - _height) > 0.0001D) {
                _cached = engine.LayoutPositionedElement(Element, Style, ParentStyle, width, height, Depth);
                _width = width;
                _height = height;
            }
            return _cached;
        }
    }

    private sealed class PositionedLayer {
        internal PositionedLayer(HtmlRenderFlowBlock block, double x, double y) {
            Block = block;
            X = x;
            Y = y;
        }
        internal HtmlRenderFlowBlock Block { get; }
        internal double X { get; }
        internal double Y { get; }
    }
}
