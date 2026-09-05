using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private const string ComponentName = "OfficeIMO.Html.Renderer";
    private readonly IHtmlDocument _document;
    private readonly HtmlComputedStyleSet _computedStyles;
    private readonly HtmlRenderOptions _options;
    private readonly HtmlDiagnosticReport _diagnostics;
    private readonly HtmlRenderStyleResolver _styleResolver;
    private HtmlGeneratedContentSet _generatedContent;
    private readonly HtmlCounterStyleRegistry _counterStyles;
    private readonly HtmlResourceSession _resources;
    private readonly HtmlCssPageRuleSet _pageRules;
    private readonly OfficeFontFaceCollection _fonts;
    private readonly HtmlRenderMetadata _metadata;
    private readonly Uri? _baseUri;
    private readonly HtmlUrlPolicy _resourceUrlPolicy;
    private readonly CancellationToken _cancellationToken;
    private IElement? _surfaceRootElement;
    private HtmlRenderBoxStyle? _surfaceRootStyle;
    private IElement? _viewportOverflowElement;
    private HtmlRenderBoxStyle? _viewportOverflowStyle;
    private int _paintOrder;
    private int _positionedSourceOrder;
    private int _nextLogicalTextOrder;
    private int _nextSemanticNodeId;
    private long _backgroundImageTileCount;
    private long _layoutOperationCount;
    private readonly List<PositionedElementRequest> _fixedPositionedElements = new List<PositionedElementRequest>();
    private readonly List<PositionedElementRequest> _rootPositionedElements = new List<PositionedElementRequest>();
    private readonly Dictionary<IElement, List<PositionedElementRequest>> _localPositionedElements = new Dictionary<IElement, List<PositionedElementRequest>>();
    private readonly Dictionary<IElement, NormalFlowPlacement> _normalFlowPlacements = new Dictionary<IElement, NormalFlowPlacement>();
    private readonly Dictionary<IElement, PositionedContainingRect> _positionedContainingRects = new Dictionary<IElement, PositionedContainingRect>();
    private readonly Dictionary<IElement, InlineContainingRect> _inlineContainingRects = new Dictionary<IElement, InlineContainingRect>();
    private readonly Dictionary<IElement, InlineStaticPosition> _inlineStaticPositions = new Dictionary<IElement, InlineStaticPosition>();
    private readonly HashSet<IElement> _inlineStackingElements = new HashSet<IElement>();
    private readonly HashSet<IElement> _suppressedEditableLayoutRegionMarkers = new HashSet<IElement>();
    private readonly Dictionary<IElement, HtmlRenderBoxStyle> _layoutStyles = new Dictionary<IElement, HtmlRenderBoxStyle>();
    private readonly Dictionary<IElement, bool> _containsInFlowFloatCache = new Dictionary<IElement, bool>();
    private readonly Dictionary<int, int> _rootStackingPaintOrders = new Dictionary<int, int>();
    private readonly Dictionary<IElement, int> _positionedSourceOrdersByElement = new Dictionary<IElement, int>();
    private readonly Dictionary<IElement, int> _semanticNodeIds = new Dictionary<IElement, int>();
    private readonly Dictionary<IElement, FlattenedSemanticBoundary> _flattenedSemanticBoundaries = new Dictionary<IElement, FlattenedSemanticBoundary>();
    private readonly Dictionary<IElement, int> _documentOrderByElement = new Dictionary<IElement, int>();
    private readonly Dictionary<IElement, int> _footnoteNumbers = new Dictionary<IElement, int>();
    private readonly Dictionary<IElement, HtmlFootnoteEntry> _footnoteEntries = new Dictionary<IElement, HtmlFootnoteEntry>();
    private readonly HashSet<string> _namedDestinationIds = new HashSet<string>(StringComparer.Ordinal);
    private HtmlFootnotePagePlan _footnotePlan = HtmlFootnotePagePlan.Empty;
    private readonly Dictionary<int, HtmlRenderBookmarkDefinition> _bookmarkDefinitions = new Dictionary<int, HtmlRenderBookmarkDefinition>();
    private readonly Dictionary<IElement, string> _staticRadioGroupKeys = new Dictionary<IElement, string>();
    private readonly Dictionary<IElement, string> _blankValueRadioGroupKeys = new Dictionary<IElement, string>();
    private readonly Dictionary<IElement, string> _mixedDisabledRadioGroupKeys = new Dictionary<IElement, string>();
    private readonly Dictionary<IElement, string> _transparentRadioGroupKeys = new Dictionary<IElement, string>();
    private readonly Dictionary<IElement, string> _backgroundImageRadioGroupKeys = new Dictionary<IElement, string>();
    private readonly Dictionary<IElement, string> _staticRepeatedControlGroupKeys = new Dictionary<IElement, string>();
    private readonly HashSet<string> _formFieldNames = new HashSet<string>(StringComparer.Ordinal);
    private readonly Dictionary<IElement, string> _formFieldNamesByElement = new Dictionary<IElement, string>();
    private readonly Dictionary<string, string> _radioFieldNames = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly HashSet<IElement> _registeredFixedElements = new HashSet<IElement>();
    private readonly HashSet<IElement> _registeredAbsoluteElements = new HashSet<IElement>();
    private readonly HashSet<IElement> _reportedPositionStaticAnchorFallbacks = new HashSet<IElement>();
    private readonly HashSet<IElement> _reportedFloatValueFallbacks = new HashSet<IElement>();
    private readonly HashSet<IElement> _reportedOverflowValueFallbacks = new HashSet<IElement>();
    private readonly HashSet<IElement> _reportedOverflowClipMarginFallbacks = new HashSet<IElement>();
    private readonly HashSet<IElement> _reportedOverflowScrollSnapshots = new HashSet<IElement>();
    private readonly HashSet<string> _reportedBorderRadiusFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedBoxShadowFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedBorderPaintFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedOutlinePaintFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedReplacedElementFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedTransformedFormFieldFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedNonUniformFormFieldRadiusFallbacks = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedStaticRadioGroups = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedStaticRepeatedControlGroups = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedStickySources = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<IElement> _reportedComplexTextShapingElements = new HashSet<IElement>();
    private readonly HtmlShapedTextMeasurementCache _shapedTextMeasurementCache = new HtmlShapedTextMeasurementCache();
    private readonly HashSet<string> _reportedMixedRunningElementMarginBoxes = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<string> _reportedPageContinuationReflow = new HashSet<string>(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _runningStringValues = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly Dictionary<int, HtmlCssRunningElementSnapshot> _runningElementSnapshots = new Dictionary<int, HtmlCssRunningElementSnapshot>();
    private readonly List<HtmlCssRunningStringAssignment> _currentPageRunningStringAssignments = new List<HtmlCssRunningStringAssignment>();
    private int _nextRunningElementSnapshotId;
    private HtmlCssRunningStringPageContext? _currentRunningStringPage;
    private HtmlCssPageGeometry _activePageGeometry;
    private IElement? _activeSubgridOwner;
    private IReadOnlyList<double>? _activeSubgridColumnSizes;
    private IReadOnlyDictionary<string, int>? _activeSubgridColumnLineNames;
    private double _activeSubgridColumnGap;
    private IReadOnlyList<double>? _activeSubgridRowSizes;
    private IReadOnlyDictionary<string, int>? _activeSubgridRowLineNames;
    private int? _activeSubgridRowLineCount;
    private double _activeSubgridRowGap;

    internal HtmlRenderLayoutEngine(IHtmlDocument document, HtmlComputedStyleSet computedStyles, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlResourceSession? resources = null, HtmlCssPageRuleSet? pageRules = null, OfficeFontFaceCollection? fonts = null, CancellationToken cancellationToken = default) {
        _cancellationToken = cancellationToken;
        _cancellationToken.ThrowIfCancellationRequested();
        _document = document;
        _computedStyles = computedStyles;
        int documentOrder = 0;
        foreach (IElement element in document.QuerySelectorAll("*")) {
            _documentOrderByElement[element] = documentOrder++;
            if (computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? elementStyle)
                && string.Equals(elementStyle.GetValue("float").Trim(), "footnote", StringComparison.OrdinalIgnoreCase)) {
                _footnoteNumbers[element] = _footnoteNumbers.Count + 1;
            }
        }
        _options = options;
        _diagnostics = diagnostics;
        _styleResolver = new HtmlRenderStyleResolver(computedStyles, options, diagnostics);
        _counterStyles = HtmlCounterStyleRegistry.Parse(document, options);
        _generatedContent = HtmlGeneratedContentResolver.Resolve(document, computedStyles, diagnostics, options.MaxLayoutDepth, _counterStyles);
        foreach (string id in _generatedContent.TargetPageIds) _namedDestinationIds.Add(id);
        foreach (IElement linkElement in document.QuerySelectorAll("[href]")) {
            string? href = linkElement.GetAttribute("href")?.Trim();
            if (href == null || href.Length <= 1 || href[0] != '#') continue;
            string fragment = href.Substring(1);
            _namedDestinationIds.Add(string.Equals(fragment, "top", StringComparison.OrdinalIgnoreCase) ? "top" : fragment);
        }
        _resources = resources ?? new HtmlResourceSession();
        _pageRules = pageRules ?? new HtmlCssPageRuleSet();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        _metadata = HtmlRenderMetadata.FromDocument(document, ResolveDocumentDirection(document, computedStyles));
        _baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        _resourceUrlPolicy = HtmlResourceUrlPolicy.Create(options.GetResourceUrlPolicy());
        _activePageGeometry = new HtmlCssPageGeometry(options.PageWidth, options.PageHeight, options.Margins);
    }

    private bool TryResolveLength(string? value, double reference, double fontSize, out double result) =>
        HtmlRenderCssValues.TryLength(
            value,
            reference,
            fontSize,
            _options.DefaultFontSize,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
            out result);

    private void SetActivePageGeometry(HtmlCssPageGeometry geometry) {
        _activePageGeometry = geometry;
        _styleResolver.SetViewport(geometry.Width, geometry.Height);
    }

    private double ActiveSurfaceWidth => _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth;
    private HtmlRenderMargins ActiveMargins => _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Margins : _options.Margins;

    private static HtmlRenderTextDirection ResolveDocumentDirection(IHtmlDocument document, HtmlComputedStyleSet computedStyles) {
        IElement? root = document.DocumentElement;
        if (root != null && computedStyles.Elements.TryGetValue(root, out HtmlComputedStyle? style)) {
            string computedDirection = style.GetValue("direction").Trim();
            if (string.Equals(computedDirection, "rtl", StringComparison.OrdinalIgnoreCase)) {
                return HtmlRenderTextDirection.RightToLeft;
            }
            if (string.Equals(computedDirection, "ltr", StringComparison.OrdinalIgnoreCase)) {
                return HtmlRenderTextDirection.LeftToRight;
            }
            // A computed root style is authoritative even when CSS reset the
            // presentational dir hint to its initial (LTR) value.
            return HtmlRenderTextDirection.LeftToRight;
        }

        string? attributeDirection = root?.GetAttribute("dir");
        return string.Equals(attributeDirection?.Trim(), "rtl", StringComparison.OrdinalIgnoreCase)
            ? HtmlRenderTextDirection.RightToLeft
            : HtmlRenderTextDirection.LeftToRight;
    }

    internal HtmlRenderDocument Render() {
        CheckCancellation();
        IdentifyStaticRadioGroups();
        HtmlRenderDocument rendered = RenderSinglePass();
        if (_options.Mode != HtmlRenderMode.Paged) return rendered;

        for (int pass = 0; pass < 4; pass++) {
            bool changed = false;
            if (_generatedContent.HasTargetPageReferences) {
                IReadOnlyDictionary<string, int> targetPages = ResolveTargetPages(rendered);
                if (!_generatedContent.TargetPagesEqual(targetPages)) {
                    _generatedContent = _generatedContent.WithTargetPages(targetPages);
                    changed = true;
                }
            }
            HtmlFootnotePagePlan footnotePlan = ResolveFootnotePlan(rendered);
            if (!_footnotePlan.EquivalentTo(footnotePlan)) {
                _footnotePlan = footnotePlan;
                changed = true;
            }
            if (!changed) break;
            ResetLayoutPassState();
            rendered = RenderSinglePass();
        }
        return rendered;
    }

    private HtmlRenderDocument RenderSinglePass() {
        CheckCancellation();
        IElement root = _document.Body ?? _document.DocumentElement ?? throw new InvalidOperationException("The parsed HTML document has no renderable root element.");
        HtmlCssPageGeometry initialGeometry = _options.Mode == HtmlRenderMode.Paged
            ? _pageRules.ResolveGeometry(1, null, _options)
            : new HtmlCssPageGeometry(_options.ViewportWidth, _options.ViewportHeight ?? 1056D, _options.Margins);
        SetActivePageGeometry(initialGeometry);
        double surfaceWidth = initialGeometry.Width;
        double contentWidth = initialGeometry.ContentWidth;
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, contentWidth);
        _layoutStyles[root] = rootStyle.Clone();
        _surfaceRootElement = root;
        _surfaceRootStyle = rootStyle;
        _viewportOverflowElement = root;
        _viewportOverflowStyle = rootStyle;
        IElement? documentRoot = _document.DocumentElement;
        if (documentRoot != null && !ReferenceEquals(documentRoot, root)) {
            HtmlRenderBoxStyle documentRootStyle = _styleResolver.Resolve(documentRoot, contentWidth);
            if (HasDeclaredCanvasBackground(documentRootStyle)) {
                _surfaceRootElement = documentRoot;
                _surfaceRootStyle = documentRootStyle;
            }
            if (HasNonVisibleOverflow(documentRootStyle)) {
                _viewportOverflowElement = documentRoot;
                _viewportOverflowStyle = documentRootStyle;
            }
        }

        IReadOnlyList<HtmlRenderFlowBlock> blocks = rootStyle.Display == "none"
            ? Array.Empty<HtmlRenderFlowBlock>()
            : BuildChildBlocks(root, contentWidth, rootStyle, 0);
        blocks = AddDocumentTopDestination(blocks, contentWidth);
        if (_options.Mode == HtmlRenderMode.Paged && blocks.Count > 0 && blocks[0].PageName != null) {
            HtmlCssPageGeometry namedGeometry = _pageRules.ResolveGeometry(1, blocks[0].PageName, _options);
            if (!SamePageGeometry(initialGeometry, namedGeometry)) {
                ResetLayoutPassState();
                SetActivePageGeometry(namedGeometry);
                contentWidth = namedGeometry.ContentWidth;
                rootStyle = _styleResolver.Resolve(root, contentWidth);
                _layoutStyles[root] = rootStyle.Clone();
                _surfaceRootStyle = rootStyle;
                blocks = rootStyle.Display == "none"
                    ? Array.Empty<HtmlRenderFlowBlock>()
                    : BuildChildBlocks(root, contentWidth, rootStyle, 0);
                blocks = AddDocumentTopDestination(blocks, contentWidth);
            }
        }
        HtmlRenderDocument rendered = _options.Mode == HtmlRenderMode.Paged
            ? RenderPaged(blocks)
            : RenderContinuous(blocks);
        CheckCancellation();
        return rendered;
    }

    private IReadOnlyList<HtmlRenderFlowBlock> AddDocumentTopDestination(
        IReadOnlyList<HtmlRenderFlowBlock> blocks,
        double contentWidth) {
        if (!_namedDestinationIds.Contains("top")) return blocks;
        var destination = new HtmlRenderNamedDestination("top", 0D, 0D, 0, "document:top-anchor");
        if (blocks.Count == 0) {
            return new[] {
                new HtmlRenderFlowBlock(
                    contentWidth,
                    0.01D,
                    new HtmlRenderVisual[] { destination },
                    HtmlPageBreakTarget.None,
                    HtmlPageBreakTarget.None,
                    false,
                    "document:top-anchor")
            };
        }

        var result = blocks.ToList();
        HtmlRenderFlowBlock first = result[0];
        result[0] = first.WithVisuals(new HtmlRenderVisual[] { destination }
            .Concat(first.Visuals.Select((visual, index) => visual.Translate(0D, 0D, index + 1))));
        return result;
    }

    private IReadOnlyDictionary<string, int> ResolveTargetPages(HtmlRenderDocument rendered) {
        var targetPages = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string id in _generatedContent.TargetPageIds) {
            IElement? target = _document.GetElementById(id);
            if (target == null) continue;
            string targetSource = HtmlRenderStyleResolver.DescribeSource(target);
            string anchorSource = HtmlRenderStyleResolver.DescribeSource(target) + ":target-page-anchor";
            foreach (HtmlRenderPage page in rendered.Pages) {
                if (!ContainsVisualSource(page.Scene, targetSource)) continue;
                targetPages[id] = page.PageNumber;
                break;
            }
            if (targetPages.ContainsKey(id)) continue;
            foreach (HtmlRenderPage page in rendered.Pages) {
                if (!ContainsVisualSource(page.Scene, anchorSource)) continue;
                targetPages[id] = page.PageNumber;
                break;
            }
        }
        return targetPages;
    }

    private static bool ContainsVisualSource(IEnumerable<HtmlRenderVisual> visuals, string source) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (string.Equals(visual.Source, source, StringComparison.Ordinal)) return true;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderLayoutRegion region ? region.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals
                : null;
            if (children != null && ContainsVisualSource(children, source)) return true;
        }
        return false;
    }

    private static bool SamePageGeometry(HtmlCssPageGeometry left, HtmlCssPageGeometry right) =>
        Math.Abs(left.Width - right.Width) <= 0.0001D
        && Math.Abs(left.Height - right.Height) <= 0.0001D
        && Math.Abs(left.Margins.Left - right.Margins.Left) <= 0.0001D
        && Math.Abs(left.Margins.Top - right.Margins.Top) <= 0.0001D
        && Math.Abs(left.Margins.Right - right.Margins.Right) <= 0.0001D
        && Math.Abs(left.Margins.Bottom - right.Margins.Bottom) <= 0.0001D;

    private void ResetLayoutPassState() {
        _surfaceRootElement = null;
        _surfaceRootStyle = null;
        _viewportOverflowElement = null;
        _viewportOverflowStyle = null;
        _paintOrder = 0;
        _positionedSourceOrder = 0;
        _nextLogicalTextOrder = 0;
        _nextSemanticNodeId = 0;
        _backgroundImageTileCount = 0;
        _layoutOperationCount = 0;
        _fixedPositionedElements.Clear();
        _rootPositionedElements.Clear();
        _localPositionedElements.Clear();
        _normalFlowPlacements.Clear();
        _positionedContainingRects.Clear();
        _inlineContainingRects.Clear();
        _inlineStaticPositions.Clear();
        _inlineStackingElements.Clear();
        _suppressedEditableLayoutRegionMarkers.Clear();
        _layoutStyles.Clear();
        _containsInFlowFloatCache.Clear();
        _rootStackingPaintOrders.Clear();
        _positionedSourceOrdersByElement.Clear();
        _semanticNodeIds.Clear();
        _flattenedSemanticBoundaries.Clear();
        _footnoteEntries.Clear();
        _bookmarkDefinitions.Clear();
        _runningStringValues.Clear();
        _runningElementSnapshots.Clear();
        _currentPageRunningStringAssignments.Clear();
        _nextRunningElementSnapshotId = 0;
        _currentRunningStringPage = null;
        _formFieldNames.Clear();
        _formFieldNamesByElement.Clear();
        _radioFieldNames.Clear();
        _registeredFixedElements.Clear();
        _registeredAbsoluteElements.Clear();
        _reportedPositionStaticAnchorFallbacks.Clear();
        _activeSubgridOwner = null;
        _activeSubgridColumnSizes = null;
        _activeSubgridColumnLineNames = null;
        _activeSubgridColumnGap = 0D;
        _activeSubgridRowSizes = null;
        _activeSubgridRowLineNames = null;
        _activeSubgridRowLineCount = null;
        _activeSubgridRowGap = 0D;
    }

    private void CheckCancellation() => _cancellationToken.ThrowIfCancellationRequested();

    private int GetSemanticNodeId(IElement element) {
        if (_semanticNodeIds.TryGetValue(element, out int nodeId)) return nodeId;
        nodeId = ++_nextSemanticNodeId;
        _semanticNodeIds[element] = nodeId;
        return nodeId;
    }

    private int GetDocumentOrder(IElement element) =>
        _documentOrderByElement.TryGetValue(element, out int order) ? order : int.MaxValue;

    private HtmlRenderDocument RenderContinuous(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        double width = _options.ViewportWidth;
        double y = _options.Margins.Top;
        var placements = new List<FlowPaintLayer>(blocks.Count);
        foreach (HtmlRenderFlowBlock block in blocks) {
            CheckCancellation();
            placements.Add(new FlowPaintLayer(block, _options.Margins.Left, y, placements.Count));
            y += block.Height;
        }

        double height = y + _options.Margins.Bottom;
        if (_options.ViewportHeight.HasValue) height = Math.Max(height, _options.ViewportHeight.Value);
        height = Math.Max(1D, height);
        ValidateSurface(width, height);

        List<HtmlRenderVisual> visuals = CreatePageVisuals(width, height);
        double contentWidth = Math.Max(1D, width - _options.Margins.Left - _options.Margins.Right);
        double contentHeight = Math.Max(1D, height - _options.Margins.Top - _options.Margins.Bottom);
        PrepareGlobalPositionedRequests(includeRoot: true, width, height, contentWidth, contentHeight);
        BuildRootStackingPaintOrders(blocks);
        AppendGlobalPositionedRequests(visuals, includeRoot: true, width, height, contentWidth, contentHeight, PositionedPaintBand.Negative);
        foreach (FlowPaintLayer placement in placements) {
            CheckCancellation();
            AddTranslatedVisuals(visuals, placement.Block.Visuals, placement.X, placement.Y, placement.Block);
        }
        AppendGlobalPositionedRequests(visuals, includeRoot: true, width, height, contentWidth, contentHeight, PositionedPaintBand.NonNegative);
        ApplyViewportOverflow(visuals, width, height);
        var page = new HtmlRenderPage(1, width, height, visuals, fonts: _fonts);
        return new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, _diagnostics, _fonts, _metadata, _bookmarkDefinitions);
    }

    private HtmlRenderDocument RenderPaged(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        _runningStringValues.Clear();
        _currentPageRunningStringAssignments.Clear();
        _currentRunningStringPage = new HtmlCssRunningStringPageContext(_runningStringValues);
        string? currentPageName = blocks.Count > 0 ? blocks[0].PageName : null;
        HtmlCssPageGeometry pageGeometry = _pageRules.ResolveGeometry(1, currentPageName, _options);
        SetActivePageGeometry(pageGeometry);
        double pageWidth = pageGeometry.Width;
        double pageHeight = pageGeometry.Height;
        double contentHeight = ResolvePageBodyContentHeight(1, pageGeometry);
        ValidateSurface(pageWidth, pageHeight);
        PrepareGlobalPositionedRequests(
            includeRoot: true,
            pageWidth,
            pageHeight,
            pageGeometry.ContentWidth,
            Math.Max(1D, contentHeight));
        BuildRootStackingPaintOrders(blocks);

        var pages = new List<HtmlRenderPage>();
        var visuals = CreatePageVisuals(pageWidth, pageHeight);
        double y = pageGeometry.Margins.Top;
        void BeginPage(string? pageName) {
            pageGeometry = _pageRules.ResolveGeometry(pages.Count + 1, pageName, _options);
            SetActivePageGeometry(pageGeometry);
            pageWidth = pageGeometry.Width;
            pageHeight = pageGeometry.Height;
            contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
            ValidateSurface(pageWidth, pageHeight);
            visuals = CreatePageVisuals(pageWidth, pageHeight);
            y = pageGeometry.Margins.Top;
        }
        for (int index = 0; index < blocks.Count; index++) {
            CheckCancellation();
            HtmlRenderFlowBlock block = blocks[index];
            bool hasPageContent = y > pageGeometry.Margins.Top + 0.0001D;
            if (!string.Equals(currentPageName, block.PageName, StringComparison.Ordinal)) {
                if (hasPageContent) CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(block.PageName);
                block = RelayoutTopLevelBlockForPage(block, pageGeometry);
                hasPageContent = false;
            } else {
                block = RelayoutTopLevelBlockForPage(block, pageGeometry);
            }

            if (!hasPageContent) currentPageName = block.PageName;
            HtmlPageBreakTarget breakBefore = ResolveForcedBreakAt(block.ForcedBreaks, 0D);
            if (breakBefore == HtmlPageBreakTarget.None) breakBefore = block.BreakBefore;
            if (breakBefore != HtmlPageBreakTarget.None) {
                ApplyBreakBefore(breakBefore, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                pageWidth = pageGeometry.Width;
                pageHeight = pageGeometry.Height;
                contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                hasPageContent = y > pageGeometry.Margins.Top + 0.0001D;
                currentPageName = block.PageName;
                block = RelayoutTopLevelBlockForPage(block, pageGeometry);
            }

            if (block.Height <= contentHeight
                && hasPageContent
                && !HasInternalForcedBreak(block)
                && y + block.Height > ResolvePageBodyBottom(pages.Count + 1, pageGeometry)) {
                CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(block.PageName);
                currentPageName = block.PageName;
                block = RelayoutTopLevelBlockForPage(block, pageGeometry);
            }

            if (block.Height <= ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y && !HasInternalForcedBreak(block)) {
                AddTranslatedVisuals(visuals, block.Visuals, pageGeometry.Margins.Left, y, block);
                RecordRunningStringAssignments(block, 0D, block.Height, y);
                y += block.Height;
            } else {
                double blockOffset = 0D;
                while (blockOffset < block.Height - 0.0001D) {
                    CheckCancellation();
                    HtmlRenderContinuationGroup? continuationGroup = block.ContinuationGroups.FirstOrDefault(group => group.AppliesAt(blockOffset));
                    bool repeatContinuation = blockOffset > 0.0001D && continuationGroup != null && continuationGroup.Visuals.Count > 0 && continuationGroup.Height > 0D;
                    double continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                    double rawAvailable = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
                    HtmlRenderTrailingGroup? trailingGroup = ResolveTrailingGroup(block, blockOffset, Math.Max(0D, rawAvailable - continuationHeight), out double fragmentLimit);
                    bool repeatTrailing = trailingGroup != null && trailingGroup.Visuals.Count > 0 && trailingGroup.Height > 0D;
                    double trailingHeight = repeatTrailing ? trailingGroup!.Height : 0D;
                    double available = rawAvailable - continuationHeight - trailingHeight;
                    bool forcedBreakFits = TryGetNextForcedBreak(block.ForcedBreaks, blockOffset, out HtmlRenderForcedBreak? forcedBreak)
                        && forcedBreak!.Offset <= fragmentLimit + 0.0001D
                        && forcedBreak.Offset <= blockOffset + available + 0.0001D;
                    double fragmentEnd = forcedBreakFits
                        ? forcedBreak!.Offset
                        : available > 0.0001D
                            ? FindFragmentEnd(block, blockOffset, available, fragmentLimit)
                            : blockOffset;
                    if (fragmentEnd <= blockOffset + 0.0001D) {
                        if (y > pageGeometry.Margins.Top + 0.0001D) {
                            CommitPage(pages, visuals, pageGeometry, currentPageName);
                            BeginPage(currentPageName);
                            continue;
                        }

                        bool originalContinuation = repeatContinuation;
                        bool originalTrailing = repeatTrailing;
                        bool foundFallback = false;
                        if (originalContinuation) {
                            double candidateAvailable = rawAvailable - trailingHeight;
                            double candidateEnd = candidateAvailable > 0.0001D
                                ? FindFragmentEnd(block, blockOffset, candidateAvailable, fragmentLimit)
                                : blockOffset;
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = false;
                                continuationHeight = 0D;
                                available = candidateAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (!foundFallback && originalTrailing) {
                            double candidateAvailable = rawAvailable - (originalContinuation ? continuationGroup!.Height : 0D);
                            double candidateEnd = candidateAvailable > 0.0001D
                                ? FindFragmentEnd(block, blockOffset, candidateAvailable, fragmentLimit)
                                : blockOffset;
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = originalContinuation;
                                continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                                repeatTrailing = false;
                                trailingHeight = 0D;
                                available = candidateAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (!foundFallback && originalContinuation && originalTrailing) {
                            double candidateEnd = FindFragmentEnd(block, blockOffset, rawAvailable, fragmentLimit);
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = false;
                                continuationHeight = 0D;
                                repeatTrailing = false;
                                trailingHeight = 0D;
                                available = rawAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (foundFallback) {
                            if (originalContinuation && !repeatContinuation) {
                                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "A repeated table header was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
                            }

                            if (originalTrailing && !repeatTrailing) {
                                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed, "A repeated table footer was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
                            }
                        } else {
                            repeatContinuation = false;
                            continuationHeight = 0D;
                            repeatTrailing = false;
                            trailingHeight = 0D;
                            available = Math.Max(0D, rawAvailable);
                            fragmentEnd = Math.Min(fragmentLimit, blockOffset + available);
                            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ForcedFragment, "A layout block had no safe break opportunity within one page and was force-fragmented.", HtmlDiagnosticSeverity.Warning, block.Source);
                        }
                    }

                    if (repeatContinuation) {
                        AddTranslatedVisuals(visuals, continuationGroup!.Visuals, pageGeometry.Margins.Left, y, block);
                        y += continuationHeight;
                    }

                    IReadOnlyList<HtmlRenderVisual> fragment = SliceBlockVisuals(block, blockOffset, fragmentEnd);
                    AddTranslatedVisuals(visuals, fragment, pageGeometry.Margins.Left, y, block);
                    RecordRunningStringAssignments(block, blockOffset, fragmentEnd, y);
                    y += fragmentEnd - blockOffset;
                    blockOffset = fragmentEnd;
                    if (repeatTrailing) {
                        AddTranslatedVisuals(visuals, trailingGroup!.Visuals, pageGeometry.Margins.Left, y, block);
                        y += trailingHeight;
                        if (blockOffset >= trailingGroup.ContentEndsAt - 0.0001D) blockOffset = trailingGroup.SourceEndsAt;
                    }

                    if (blockOffset < block.Height - 0.0001D) {
                        HtmlInlineBreakProgress? continuationProgress = ResolveInlineContinuationProgress(block, blockOffset);
                        string? nextPageName = ResolvePageNameAt(block.ForcedBreaks, blockOffset, currentPageName);
                        CommitPage(pages, visuals, pageGeometry, currentPageName);
                        BeginPage(nextPageName);
                        currentPageName = nextPageName;
                        pageWidth = pageGeometry.Width;
                        pageHeight = pageGeometry.Height;
                        contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                        HtmlPageBreakTarget internalBreak = ResolveForcedBreakAt(block.ForcedBreaks, blockOffset);
                        if (internalBreak != HtmlPageBreakTarget.None) {
                            EnsurePageSide(internalBreak, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                            pageWidth = pageGeometry.Width;
                            pageHeight = pageGeometry.Height;
                            contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                        }
                        if (RequiresPageRelayout(block, pageGeometry)) {
                            if (continuationProgress.HasValue
                                && TryRelayoutInlineContinuation(block, pageGeometry, continuationProgress.Value, out HtmlRenderFlowBlock reflowed)) {
                                block = reflowed;
                                blockOffset = 0D;
                            } else {
                                ReportPageContinuationReflowPending(block, pageGeometry);
                            }
                        }
                    }
                }
            }

            HtmlPageBreakTarget breakAfter = ResolveForcedBreakAt(block.ForcedBreaks, block.Height);
            if (block.BreakAfter != HtmlPageBreakTarget.None) breakAfter = block.BreakAfter;
            if (breakAfter != HtmlPageBreakTarget.None && index < blocks.Count - 1) {
                CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(currentPageName);
                EnsurePageSide(breakAfter, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                pageWidth = pageGeometry.Width;
                pageHeight = pageGeometry.Height;
                contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
            }
        }

        CommitPage(pages, visuals, pageGeometry, currentPageName);
        while (pages.Count < _footnotePlan.MaximumPageNumber) {
            BeginPage(currentPageName);
            CommitPage(pages, visuals, pageGeometry, currentPageName);
        }
        return new HtmlRenderDocument(HtmlRenderMode.Paged, ApplyPageMarginContent(pages), _diagnostics, _fonts, _metadata, _bookmarkDefinitions);
    }

    private HtmlRenderFlowBlock RelayoutTopLevelBlockForPage(HtmlRenderFlowBlock block, HtmlCssPageGeometry geometry) {
        if (!RequiresPageRelayout(block, geometry)) return block;
        if (block.OwnerElement == null) {
            ReportPageContinuationReflowPending(block, geometry);
            return block;
        }
        IElement root = _document.Body ?? _document.DocumentElement ?? block.OwnerElement;
        if (!ReferenceEquals(block.OwnerElement.ParentElement, root)) {
            ReportPageContinuationReflowPending(block, geometry);
            return block;
        }
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = _styleResolver.Resolve(block.OwnerElement, geometry.ContentWidth, rootStyle);
        return LayoutElement(block.OwnerElement, geometry.ContentWidth, style, rootStyle, 1);
    }

    private static bool RequiresPageRelayout(HtmlRenderFlowBlock block, HtmlCssPageGeometry geometry) =>
        !IsGeometryIndependentAssignmentMarker(block)
        && (Math.Abs(block.Width - geometry.ContentWidth) > 0.0001D
        || double.IsNaN(block.LayoutViewportWidth)
        || double.IsNaN(block.LayoutViewportHeight)
        || Math.Abs(block.LayoutViewportWidth - geometry.Width) > 0.0001D
        || Math.Abs(block.LayoutViewportHeight - geometry.Height) > 0.0001D);

    private static bool IsGeometryIndependentAssignmentMarker(HtmlRenderFlowBlock block) =>
        block.Height <= 0.0001D
        && block.Visuals.Count == 0
        && block.RunningStringAssignments.Count > 0;

    private void ReportPageContinuationReflowPending(HtmlRenderFlowBlock block, HtmlCssPageGeometry geometry) {
        string key = block.Source
            + "|" + geometry.ContentWidth.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            + "|" + geometry.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            + "|" + geometry.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        if (!_reportedPageContinuationReflow.Add(key)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.PagePseudoGeometryPending,
            "A fragment crossing into a page with different geometry retained its source-page layout.",
            HtmlDiagnosticSeverity.Warning,
            block.Source,
            "target-content-width=" + geometry.ContentWidth.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture));
    }

    private static HtmlInlineBreakProgress? ResolveInlineContinuationProgress(HtmlRenderFlowBlock block, double fragmentEnd) {
        if (!block.SupportsInlineContinuationReflow) return null;
        HtmlInlineBreakProgress? selected = null;
        foreach (HtmlInlineBreakProgress item in block.InlineBreakProgress) {
            if (item.OwnerElement != null && Math.Abs(item.Offset - fragmentEnd) <= 0.0001D) selected = item;
        }
        return selected;
    }

    private bool TryRelayoutInlineContinuation(
        HtmlRenderFlowBlock source,
        HtmlCssPageGeometry geometry,
        HtmlInlineBreakProgress continuation,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = source;
        if (source.OwnerElement == null || continuation.OwnerElement == null) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? source.OwnerElement;
        if (!ReferenceEquals(source.OwnerElement.ParentElement, root)) return false;
        if (!ContainsElementOrSelf(source.OwnerElement, continuation.OwnerElement)) return false;
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = _styleResolver.Resolve(source.OwnerElement, geometry.ContentWidth, rootStyle);
        reflowed = LayoutElement(
            source.OwnerElement,
            geometry.ContentWidth,
            style,
            rootStyle,
            1,
            continuation.OwnerElement,
            continuation.LogicalCharacters);
        return true;
    }

    private List<HtmlRenderVisual> CreatePageVisuals(double width, double height) {
        var visuals = new List<HtmlRenderVisual> { CreatePageBackground(width, height) };
        if (_surfaceRootElement == null || _surfaceRootStyle == null || !_surfaceRootStyle.PaintVisible || _surfaceRootStyle.Display == "none") return visuals;

        var rootBackground = new List<HtmlRenderVisual>();
        AddBoxBackground(
            rootBackground,
            _surfaceRootStyle,
            0D,
            0D,
            width,
            height,
            0D,
            _surfaceRootElement,
            HtmlRenderStyleResolver.DescribeSource(_surfaceRootElement),
            "render-root-background");
        for (int index = 0; index < rootBackground.Count; index++) {
            visuals.Add(rootBackground[index].Translate(0D, 0D, int.MinValue + 1 + index));
        }

        return visuals;
    }

    private void ApplyBreakBefore(
        HtmlPageBreakTarget target,
        ICollection<HtmlRenderPage> pages,
        ref List<HtmlRenderVisual> visuals,
        ref double y,
        ref HtmlCssPageGeometry geometry,
        string? pageName) {
        if (y > geometry.Margins.Top + 0.0001D) {
            CommitPage(pages, visuals, geometry, pageName);
            geometry = _pageRules.ResolveGeometry(pages.Count + 1, pageName, _options);
            SetActivePageGeometry(geometry);
            ValidateSurface(geometry.Width, geometry.Height);
            visuals = CreatePageVisuals(geometry.Width, geometry.Height);
            y = geometry.Margins.Top;
        }

        EnsurePageSide(target, pages, ref visuals, ref y, ref geometry, pageName);
    }

    private void EnsurePageSide(
        HtmlPageBreakTarget target,
        ICollection<HtmlRenderPage> pages,
        ref List<HtmlRenderVisual> visuals,
        ref double y,
        ref HtmlCssPageGeometry geometry,
        string? pageName) {
        if (target != HtmlPageBreakTarget.Left && target != HtmlPageBreakTarget.Right) return;
        int nextPageNumber = pages.Count + 1;
        bool nextIsRight = nextPageNumber % 2 != 0;
        bool targetIsRight = target == HtmlPageBreakTarget.Right;
        if (nextIsRight == targetIsRight) return;
        CommitPage(pages, visuals, geometry, pageName);
        geometry = _pageRules.ResolveGeometry(pages.Count + 1, pageName, _options);
        SetActivePageGeometry(geometry);
        ValidateSurface(geometry.Width, geometry.Height);
        visuals = CreatePageVisuals(geometry.Width, geometry.Height);
        y = geometry.Margins.Top;
    }

    private HtmlRenderShape CreatePageBackground(double width, double height) {
        OfficeShape background = OfficeShape.Rectangle(width, height);
        background.FillColor = _options.BackgroundColor;
        background.StrokeWidth = 0D;
        return new HtmlRenderShape(background, 0D, 0D, int.MinValue, source: "render-surface");
    }

    private static bool HasDeclaredCanvasBackground(HtmlRenderBoxStyle style) =>
        style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0
        || style.HasDeclaredBackgroundImage;

    private void CommitPage(
        ICollection<HtmlRenderPage> pages,
        List<HtmlRenderVisual> visuals,
        HtmlCssPageGeometry geometry,
        string? pageName) {
        if (pages.Count >= _options.MaxPageCount) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
        }

        double width = geometry.Width;
        double height = geometry.Height;
        bool includeRoot = pages.Count == 0;
        double contentWidth = geometry.ContentWidth;
        double contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, geometry);
        AddFootnoteVisuals(visuals, pages.Count + 1, geometry);
        PrepareGlobalPositionedRequests(includeRoot, width, height, contentWidth, contentHeight);
        AppendGlobalPositionedRequests(visuals, includeRoot, width, height, contentWidth, contentHeight, PositionedPaintBand.Negative);
        AppendGlobalPositionedRequests(visuals, includeRoot, width, height, contentWidth, contentHeight, PositionedPaintBand.NonNegative);
        CollectGlobalPositionedRunningStringAssignments(
            _currentPageRunningStringAssignments,
            includeRoot,
            width,
            height,
            contentWidth,
            contentHeight);
        if (_currentRunningStringPage != null) {
            foreach (HtmlCssRunningStringAssignment assignment in _currentPageRunningStringAssignments.OrderBy(item => item.OrderOffset)) {
                _currentRunningStringPage.Assign(assignment, _runningStringValues);
            }
        }
        ApplyViewportOverflow(visuals, width, height);
        AddPrintProductionMarks(visuals, geometry);
        pages.Add(new HtmlRenderPage(pages.Count + 1, width, height, visuals, pageName, _fonts, _currentRunningStringPage, geometry.Margins, geometry.PrintProduction));
        _currentPageRunningStringAssignments.Clear();
        _currentRunningStringPage = new HtmlCssRunningStringPageContext(_runningStringValues);
    }

    private void AddPrintProductionMarks(List<HtmlRenderVisual> visuals, HtmlCssPageGeometry geometry) {
        HtmlRenderPrintProductionSettings? production = geometry.PrintProduction;
        if (production == null || production.Marks == HtmlRenderPrintMarks.None) return;

        double trimInset = production.TrimInset;
        double trimRight = geometry.Width - trimInset;
        double trimBottom = geometry.Height - trimInset;
        double markGap = 4D;
        double markLength = 12D;
        void AddLine(double x1, double y1, double x2, double y2, string suffix) {
            double width = Math.Max(0.0001D, x2 - x1);
            double height = Math.Max(0.0001D, y2 - y1);
            OfficeShape line = OfficeShape.Line(0D, 0D, width, height);
            line.FillColor = null;
            line.StrokeColor = OfficeColor.Black;
            line.StrokeWidth = 0.6666666667D;
            visuals.Add(new HtmlRenderShape(line, x1, y1, _paintOrder++, source: "@page marks:" + suffix));
        }

        if ((production.Marks & HtmlRenderPrintMarks.Crop) != 0) {
            AddLine(trimInset - markGap - markLength, trimInset, trimInset - markGap, trimInset, "crop-top-left-h");
            AddLine(trimInset, trimInset - markGap - markLength, trimInset, trimInset - markGap, "crop-top-left-v");
            AddLine(trimRight + markGap, trimInset, trimRight + markGap + markLength, trimInset, "crop-top-right-h");
            AddLine(trimRight, trimInset - markGap - markLength, trimRight, trimInset - markGap, "crop-top-right-v");
            AddLine(trimInset - markGap - markLength, trimBottom, trimInset - markGap, trimBottom, "crop-bottom-left-h");
            AddLine(trimInset, trimBottom + markGap, trimInset, trimBottom + markGap + markLength, "crop-bottom-left-v");
            AddLine(trimRight + markGap, trimBottom, trimRight + markGap + markLength, trimBottom, "crop-bottom-right-h");
            AddLine(trimRight, trimBottom + markGap, trimRight, trimBottom + markGap + markLength, "crop-bottom-right-v");
        }

        if ((production.Marks & HtmlRenderPrintMarks.Cross) != 0) {
            double centerX = geometry.Width / 2D;
            double centerY = geometry.Height / 2D;
            double radius = 4D;
            double topY = production.MarkArea / 2D;
            double bottomY = geometry.Height - topY;
            double leftX = production.MarkArea / 2D;
            double rightX = geometry.Width - leftX;
            AddLine(centerX - radius, topY, centerX + radius, topY, "cross-top-h");
            AddLine(centerX, topY - radius, centerX, topY + radius, "cross-top-v");
            AddLine(centerX - radius, bottomY, centerX + radius, bottomY, "cross-bottom-h");
            AddLine(centerX, bottomY - radius, centerX, bottomY + radius, "cross-bottom-v");
            AddLine(leftX - radius, centerY, leftX + radius, centerY, "cross-left-h");
            AddLine(leftX, centerY - radius, leftX, centerY + radius, "cross-left-v");
            AddLine(rightX - radius, centerY, rightX + radius, centerY, "cross-right-h");
            AddLine(rightX, centerY - radius, rightX, centerY + radius, "cross-right-v");
        }
    }

    private void RecordRunningStringAssignments(HtmlRenderFlowBlock block, double start, double end, double pageOffset) {
        if (_currentRunningStringPage == null || block.RunningStringAssignments.Count == 0) return;
        if (end < start - 0.0001D) return;
        bool finalFragment = end >= block.Height - 0.0001D;
        foreach (HtmlCssRunningStringAssignment assignment in block.RunningStringAssignments) {
            if (assignment.Offset < start - 0.0001D) continue;
            if (assignment.Offset > end + 0.0001D) continue;
            if (!finalFragment && assignment.Offset >= end - 0.0001D) continue;
            _currentPageRunningStringAssignments.Add(assignment.Translate(pageOffset - start));
        }
    }

    private void AddTranslatedVisuals(
        ICollection<HtmlRenderVisual> target,
        IEnumerable<HtmlRenderVisual> source,
        double offsetX,
        double offsetY,
        HtmlRenderFlowBlock? stackingBlock = null) {
        foreach (HtmlRenderVisual visual in source) {
            int paintOrder = stackingBlock?.StackingZIndex.HasValue == true
                ? ResolveRootStackingPaintOrder(stackingBlock.StackingSourceOrder, _paintOrder++)
                : _paintOrder++;
            target.Add(visual.Translate(offsetX, offsetY, paintOrder));
        }
    }

    private IReadOnlyList<HtmlRenderVisual> SliceBlockVisuals(HtmlRenderFlowBlock block, double start, double end) {
        return SliceVisuals(block.Visuals, start, end);
    }

    private IReadOnlyList<HtmlRenderVisual> SliceVisuals(IEnumerable<HtmlRenderVisual> sourceVisuals, double start, double end) {
        var fragment = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in sourceVisuals) {
            double visualTop = visual.LayoutY;
            double visualBottom = visual.LayoutY + visual.Height;
            double intersectionTop = Math.Max(start, visualTop);
            double intersectionBottom = Math.Min(end, visualBottom);
            if (intersectionBottom <= intersectionTop + 0.0001D) continue;

            bool fullyContained = visualTop >= start - 0.0001D && visualBottom <= end + 0.0001D;
            if (fullyContained) {
                fragment.Add(visual.Translate(0D, -start, fragment.Count));
                continue;
            }

            if (visual is HtmlRenderClipGroup clipGroup) {
                IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(clipGroup.Visuals, start, end);
                if (children.Count > 0) {
                    fragment.Add(new HtmlRenderClipGroup(
                        clipGroup.ClipX,
                        clipGroup.ClipY - start,
                        clipGroup.ClipWidth,
                        clipGroup.ClipHeight,
                        clipGroup.ClipHorizontal,
                        clipGroup.ClipVertical,
                        children,
                        fragment.Count,
                        clipGroup.Source,
                        Math.Max(start, clipGroup.LayoutY) - start));
                }
                continue;
            }

            if (visual is HtmlRenderSemanticGroup semanticGroup) {
                IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(semanticGroup.Visuals, start, end);
                if (children.Count > 0) {
                    fragment.Add(new HtmlRenderSemanticGroup(
                        semanticGroup.Role,
                        semanticGroup.X,
                        semanticGroup.Y - start,
                        semanticGroup.Width,
                        Math.Max(0.01D, intersectionBottom - intersectionTop),
                        children,
                        fragment.Count,
                        semanticGroup.Source,
                        semanticGroup.ColumnSpan,
                        semanticGroup.RowSpan,
                        semanticGroup.HeaderScope,
                        semanticGroup.LayoutY - start,
                        semanticGroup.StructureElementKey));
                }
                continue;
            }

            if (visual is HtmlRenderLayoutRegion layoutRegion) {
                IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(layoutRegion.Visuals, start, end);
                fragment.Add(new HtmlRenderLayoutRegion(
                    layoutRegion.SourceKey,
                    layoutRegion.RegionKind,
                    layoutRegion.SourceText,
                    layoutRegion.Position,
                    layoutRegion.FloatSide,
                    layoutRegion.ZIndex,
                    layoutRegion.BackgroundLayerCount,
                    layoutRegion.BoxShadowLayerCount,
                    layoutRegion.BackgroundColor,
                    layoutRegion.X,
                    layoutRegion.Y - start,
                    layoutRegion.Width,
                    Math.Max(0.01D, intersectionBottom - intersectionTop),
                    children,
                    fragment.Count,
                    layoutRegion.Source,
                    layoutRegion.LayoutY - start));
                continue;
            }

            if (visual is HtmlRenderBookmarkAnchor bookmarkAnchor) {
                if (bookmarkAnchor.LayoutY >= start - 0.0001D && bookmarkAnchor.LayoutY < end - 0.0001D) {
                    fragment.Add(new HtmlRenderBookmarkAnchor(
                        bookmarkAnchor.SemanticNodeId,
                        bookmarkAnchor.Text,
                        bookmarkAnchor.X,
                        bookmarkAnchor.Y - start,
                        bookmarkAnchor.Width,
                        bookmarkAnchor.Height,
                        fragment.Count,
                        bookmarkAnchor.Source,
                        bookmarkAnchor.LayoutY - start));
                }
                continue;
            }

            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(logicalTextGroup.Visuals, start, end);
                if (children.Count > 0) {
                    fragment.Add(new HtmlRenderLogicalTextGroup(
                        ResolveLogicalText(children, logicalTextGroup.Text),
                        logicalTextGroup.X,
                        logicalTextGroup.Y - start,
                        logicalTextGroup.Width,
                        Math.Max(0.01D, intersectionBottom - intersectionTop),
                        children,
                        fragment.Count,
                        logicalTextGroup.Source,
                        logicalTextGroup.LayoutY - start));
                }
                continue;
            }

            if (visual is HtmlRenderEffectGroup effectGroup) {
                IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(effectGroup.Visuals, start, end);
                if (children.Count > 0) {
                    double translatedY = -start;
                    OfficeTransform transform = OfficeTransform.Translate(0D, -translatedY)
                        .Then(effectGroup.Transform)
                        .Then(OfficeTransform.Translate(0D, translatedY));
                    fragment.Add(new HtmlRenderEffectGroup(
                        effectGroup.X,
                        effectGroup.Y - start,
                        effectGroup.Width,
                        Math.Max(0.01D, intersectionBottom - intersectionTop),
                        transform,
                        effectGroup.Opacity,
                        children,
                        fragment.Count,
                        effectGroup.Source,
                        Math.Max(start, effectGroup.LayoutY) - start));
                }
                continue;
            }

            if (visual is HtmlRenderImage
                || visual is HtmlRenderDrawing
                || visual is HtmlRenderImagePattern
                || visual is HtmlRenderPathClipGroup
                || visual is HtmlRenderShape) {
                fragment.Add(CreateVerticallyClippedVisualFragment(visual, start, intersectionTop, intersectionBottom, fragment.Count));
                continue;
            }

            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "A visual crossing a forced page boundary could not be represented safely in the current fragment.", HtmlDiagnosticSeverity.Warning, visual.Source, visual.Kind.ToString());
        }

        return fragment;
    }

    private static HtmlRenderClipGroup CreateVerticallyClippedVisualFragment(
        HtmlRenderVisual visual,
        double fragmentStart,
        double intersectionTop,
        double intersectionBottom,
        int paintOrder) {
        double clipY = intersectionTop - fragmentStart;
        return new HtmlRenderClipGroup(
            visual.X,
            clipY,
            visual.Width,
            Math.Max(0.01D, intersectionBottom - intersectionTop),
            clipHorizontal: false,
            clipVertical: true,
            new[] { visual.Translate(0D, -fragmentStart, 0) },
            paintOrder,
            visual.Source,
            clipY);
    }

    private void ValidateSurface(double width, double height) {
        double scale = _options.GetEffectiveScale(width, height);
        double pixelWidth = Math.Ceiling(width * scale);
        double pixelHeight = Math.Ceiling(height * scale);
        if (double.IsNaN(pixelWidth) || double.IsInfinity(pixelWidth) ||
            double.IsNaN(pixelHeight) || double.IsInfinity(pixelHeight) ||
            pixelWidth > _options.MaxSurfaceWidth || pixelHeight > _options.MaxSurfaceHeight) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum image surface dimensions.");
        }
    }

    private void AddUnsupported(
        string code,
        string message,
        IElement element,
        string? detail = null,
        OfficeConversionLossKind lossKind = OfficeConversionLossKind.Approximation) {
        _diagnostics.Add(
            ComponentName,
            code,
            message,
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            detail,
            lossKind);
    }
}
