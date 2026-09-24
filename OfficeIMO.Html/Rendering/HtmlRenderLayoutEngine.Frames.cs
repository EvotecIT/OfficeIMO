using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutFrame(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style) {
        const double intrinsicWidth = 300D;
        const double intrinsicHeight = 150D;
        ReplacedContentSize contentSize = ResolveReplacedContentSize(
            style, intrinsicWidth, intrinsicHeight, hasIntrinsicSize: true);
        double boxWidth = contentSize.Width + style.HorizontalInsets;
        double boxHeight = contentSize.Height + style.VerticalInsets;
        EnsureReplacedBoxSize(boxWidth, boxHeight);

        string source = HtmlRenderStyleResolver.DescribeSource(element);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        IReadOnlyList<HtmlRenderVisual> frameVisuals = RenderFrameContent(
            element, contentSize.Width, contentSize.Height, source);
        IReadOnlyList<HtmlRenderVisual> translated = frameVisuals
            .Select((visual, index) => visual.Translate(contentX, contentY, index))
            .ToArray();

        HtmlResolvedBorderRadii outerRadii = ResolveBoxRadii(style, boxWidth, boxHeight, element, source);
        HtmlResolvedBorderRadii contentRadii = outerRadii.Inset(
            style.BorderLeftWidth + style.PaddingLeft,
            style.BorderTopWidth + style.PaddingTop,
            style.BorderRightWidth + style.PaddingRight,
            style.BorderBottomWidth + style.PaddingBottom,
            contentSize.Width,
            contentSize.Height);
        IEnumerable<HtmlRenderVisual> clippedChildren = translated;
        HtmlResolvedBorderRadii normalized = contentRadii.Normalize(contentSize.Width, contentSize.Height);
        if (!normalized.IsZero && translated.Count > 0) {
            clippedChildren = new[] {
                new HtmlRenderPathClipGroup(
                    contentX,
                    contentY,
                    CreateBoxClipPath(contentSize.Width, contentSize.Height, normalized),
                    translated,
                    0,
                    source + ":rounded-content-clip")
            };
        }
        visuals.Add(new HtmlRenderClipGroup(
            contentX,
            contentY,
            contentSize.Width,
            contentSize.Height,
            clipHorizontal: true,
            clipVertical: true,
            clippedChildren,
            visuals.Count,
            source + ":frame-viewport"));

        ReportReplacedElementFallbacks(style, element);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        if (!style.PaintVisible) visuals.Clear();
        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        return new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            pageName: style.PageName);
    }

    private IReadOnlyList<HtmlRenderVisual> RenderFrameContent(
        IElement frame,
        double viewportWidth,
        double viewportHeight,
        string source) {
        string? srcdoc = frame.GetAttribute("srcdoc");
        if (string.IsNullOrWhiteSpace(srcdoc)) return Array.Empty<HtmlRenderVisual>();
        if (_options.FrameDepth >= _options.MaxFrameDepth) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.FrameDepthLimitExceeded,
                "Nested iframe rendering exceeded the configured frame-depth limit.",
                HtmlDiagnosticSeverity.Error,
                source,
                "limit=" + _options.MaxFrameDepth,
                OfficeConversionLossKind.Omission);
            return Array.Empty<HtmlRenderVisual>();
        }

        _cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(srcdoc!, _cancellationToken);
        HtmlRenderInputGuard.ValidateDocument(document, _options, _cancellationToken);
        var options = _options.Clone();
        options.Mode = HtmlRenderMode.Continuous;
        options.CssMediaContextOverride = _options.MediaContext;
        options.ViewportWidth = Math.Max(0.01D, viewportWidth);
        options.CssMediaWidthOverride = null;
        options.ViewportHeight = Math.Max(0.01D, viewportHeight);
        options.Margins = HtmlRenderMargins.All(0D);
        options.HonorCssPageRules = false;
        options.ClipContinuousSurfaceToViewport = false;
        options.BaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, _baseUri);
        options.ResourceResolver = null;
        options.SynchronousResourceResolver = null;
        options.AdditionalStylesheets.Clear();
        options.FrameDepth++;
        options.Validate();

        HtmlCssRuleBlockScanner.ValidateDocument(document, _limits);
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(document, _limits, options);
        HtmlRenderStylesheetApplier.Apply(document, _resources, options, _limits, cssBudget, _diagnostics);
        HtmlCssRuleBlockScanner.ValidateDocument(document, _limits);
        OfficeFontFaceCollection fonts = HtmlRenderFontFaceLoader.Load(
            document, _resources, options, _limits, _diagnostics);
        fonts.AddRange(options.Fonts);
        _fonts.AddRange(fonts);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, options, _diagnostics);
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, options, _limits);
        var engine = new HtmlRenderLayoutEngine(
            document,
            styles,
            options,
            _diagnostics,
            _resources,
            pageRules,
            fonts,
            _limits,
            _nextLogicalTextOrder,
            _nextSemanticNodeId,
            _operationBudget,
            _cancellationToken);
        HtmlRenderDocument rendered = engine.Render();
        _nextLogicalTextOrder = engine._nextLogicalTextOrder;
        _nextSemanticNodeId = engine._nextSemanticNodeId;
        return ReplaceDescendantFormFieldsForPaintEffect(
            rendered.Pages.SelectMany(page => page.Scene).ToArray(),
            "iframe-content");
    }
}
