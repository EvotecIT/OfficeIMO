using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private const string ComponentName = "OfficeIMO.Html.Renderer";
    private readonly IHtmlDocument _document;
    private readonly HtmlRenderOptions _options;
    private readonly HtmlDiagnosticReport _diagnostics;
    private readonly HtmlRenderStyleResolver _styleResolver;
    private readonly HtmlGeneratedContentSet _generatedContent;
    private readonly HtmlRenderResourceSet _resources;
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
    private int _nextSemanticNodeId;
    private long _backgroundImageTileCount;
    private readonly List<PositionedElementRequest> _fixedPositionedElements = new List<PositionedElementRequest>();
    private readonly List<PositionedElementRequest> _rootPositionedElements = new List<PositionedElementRequest>();
    private readonly Dictionary<IElement, List<PositionedElementRequest>> _localPositionedElements = new Dictionary<IElement, List<PositionedElementRequest>>();
    private readonly Dictionary<IElement, NormalFlowPlacement> _normalFlowPlacements = new Dictionary<IElement, NormalFlowPlacement>();
    private readonly Dictionary<IElement, PositionedContainingRect> _positionedContainingRects = new Dictionary<IElement, PositionedContainingRect>();
    private readonly Dictionary<IElement, InlineContainingRect> _inlineContainingRects = new Dictionary<IElement, InlineContainingRect>();
    private readonly Dictionary<IElement, InlineStaticPosition> _inlineStaticPositions = new Dictionary<IElement, InlineStaticPosition>();
    private readonly HashSet<IElement> _inlineStackingElements = new HashSet<IElement>();
    private readonly Dictionary<IElement, HtmlRenderBoxStyle> _layoutStyles = new Dictionary<IElement, HtmlRenderBoxStyle>();
    private readonly Dictionary<IElement, bool> _containsInFlowFloatCache = new Dictionary<IElement, bool>();
    private readonly Dictionary<int, int> _rootStackingPaintOrders = new Dictionary<int, int>();
    private readonly Dictionary<IElement, int> _positionedSourceOrdersByElement = new Dictionary<IElement, int>();
    private readonly Dictionary<IElement, int> _semanticNodeIds = new Dictionary<IElement, int>();
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
    private readonly HashSet<string> _reportedStickySources = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<IElement> _reportedBidiElements = new HashSet<IElement>();

    internal HtmlRenderLayoutEngine(IHtmlDocument document, HtmlComputedStyleSet computedStyles, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlRenderResourceSet? resources = null, HtmlCssPageRuleSet? pageRules = null, OfficeFontFaceCollection? fonts = null, CancellationToken cancellationToken = default) {
        _cancellationToken = cancellationToken;
        _cancellationToken.ThrowIfCancellationRequested();
        _document = document;
        _options = options;
        _diagnostics = diagnostics;
        _styleResolver = new HtmlRenderStyleResolver(computedStyles, options);
        _generatedContent = HtmlGeneratedContentResolver.Resolve(document, computedStyles, diagnostics, options.MaxLayoutDepth);
        _resources = resources ?? new HtmlRenderResourceSet();
        _pageRules = pageRules ?? new HtmlCssPageRuleSet();
        _fonts = fonts?.Clone() ?? new OfficeFontFaceCollection();
        string? language = document.DocumentElement?.GetAttribute("lang");
        if (string.IsNullOrWhiteSpace(language)) language = document.DocumentElement?.GetAttribute("xml:lang");
        _metadata = new HtmlRenderMetadata(document.Title, language, ResolveDocumentDirection(document, computedStyles));
        _baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        _resourceUrlPolicy = HtmlResourceUrlPolicy.Create(options.GetResourceUrlPolicy());
    }

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
        }

        string? attributeDirection = root?.GetAttribute("dir");
        return string.Equals(attributeDirection?.Trim(), "rtl", StringComparison.OrdinalIgnoreCase)
            ? HtmlRenderTextDirection.RightToLeft
            : HtmlRenderTextDirection.LeftToRight;
    }

    internal HtmlRenderDocument Render() {
        CheckCancellation();
        IElement root = _document.Body ?? _document.DocumentElement ?? throw new InvalidOperationException("The parsed HTML document has no renderable root element.");
        double surfaceWidth = _options.Mode == HtmlRenderMode.Paged ? _options.PageWidth : _options.ViewportWidth;
        double contentWidth = surfaceWidth - _options.Margins.Left - _options.Margins.Right;
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
        HtmlRenderDocument rendered = _options.Mode == HtmlRenderMode.Paged
            ? RenderPaged(blocks)
            : RenderContinuous(blocks);
        CheckCancellation();
        return rendered;
    }

    private void CheckCancellation() => _cancellationToken.ThrowIfCancellationRequested();

    private int GetSemanticNodeId(IElement element) {
        if (_semanticNodeIds.TryGetValue(element, out int nodeId)) return nodeId;
        nodeId = ++_nextSemanticNodeId;
        _semanticNodeIds[element] = nodeId;
        return nodeId;
    }

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
        return new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, _diagnostics, _fonts, _metadata);
    }

    private HtmlRenderDocument RenderPaged(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        double pageWidth = _options.PageWidth;
        double pageHeight = _options.PageHeight;
        double contentHeight = pageHeight - _options.Margins.Top - _options.Margins.Bottom;
        ValidateSurface(pageWidth, pageHeight);
        PrepareGlobalPositionedRequests(
            includeRoot: true,
            pageWidth,
            pageHeight,
            Math.Max(1D, pageWidth - _options.Margins.Left - _options.Margins.Right),
            Math.Max(1D, contentHeight));
        BuildRootStackingPaintOrders(blocks);

        var pages = new List<HtmlRenderPage>();
        var visuals = CreatePageVisuals(pageWidth, pageHeight);
        double y = _options.Margins.Top;
        string? currentPageName = null;
        for (int index = 0; index < blocks.Count; index++) {
            CheckCancellation();
            HtmlRenderFlowBlock block = blocks[index];
            bool hasPageContent = y > _options.Margins.Top + 0.0001D;
            if (hasPageContent && !string.Equals(currentPageName, block.PageName, StringComparison.OrdinalIgnoreCase)) {
                CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
                hasPageContent = false;
            }

            if (!hasPageContent) currentPageName = block.PageName;
            if (block.BreakBefore != HtmlPageBreakTarget.None) {
                ApplyBreakBefore(block.BreakBefore, pages, ref visuals, ref y, pageWidth, pageHeight, currentPageName);
                hasPageContent = y > _options.Margins.Top + 0.0001D;
                currentPageName = block.PageName;
            }

            if (block.Height <= contentHeight && hasPageContent && y + block.Height > pageHeight - _options.Margins.Bottom) {
                CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
                currentPageName = block.PageName;
            }

            if (block.Height <= pageHeight - _options.Margins.Bottom - y) {
                AddTranslatedVisuals(visuals, block.Visuals, _options.Margins.Left, y, block);
                y += block.Height;
            } else {
                double blockOffset = 0D;
                while (blockOffset < block.Height - 0.0001D) {
                    CheckCancellation();
                    HtmlRenderContinuationGroup? continuationGroup = block.ContinuationGroups.FirstOrDefault(group => group.AppliesAt(blockOffset));
                    bool repeatContinuation = blockOffset > 0.0001D && continuationGroup != null && continuationGroup.Visuals.Count > 0 && continuationGroup.Height > 0D;
                    double continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                    double rawAvailable = pageHeight - _options.Margins.Bottom - y;
                    HtmlRenderTrailingGroup? trailingGroup = ResolveTrailingGroup(block, blockOffset, Math.Max(0D, rawAvailable - continuationHeight), out double fragmentLimit);
                    bool repeatTrailing = trailingGroup != null && trailingGroup.Visuals.Count > 0 && trailingGroup.Height > 0D;
                    double trailingHeight = repeatTrailing ? trailingGroup!.Height : 0D;
                    double available = rawAvailable - continuationHeight - trailingHeight;
                    double fragmentEnd = available > 0.0001D
                        ? FindFragmentEnd(block, blockOffset, available, fragmentLimit)
                        : blockOffset;
                    if (fragmentEnd <= blockOffset + 0.0001D) {
                        if (y > _options.Margins.Top + 0.0001D) {
                            CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
                            visuals = CreatePageVisuals(pageWidth, pageHeight);
                            y = _options.Margins.Top;
                            currentPageName = block.PageName;
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
                        AddTranslatedVisuals(visuals, continuationGroup!.Visuals, _options.Margins.Left, y, block);
                        y += continuationHeight;
                    }

                    IReadOnlyList<HtmlRenderVisual> fragment = SliceBlockVisuals(block, blockOffset, fragmentEnd);
                    AddTranslatedVisuals(visuals, fragment, _options.Margins.Left, y, block);
                    y += fragmentEnd - blockOffset;
                    blockOffset = fragmentEnd;
                    if (repeatTrailing) {
                        AddTranslatedVisuals(visuals, trailingGroup!.Visuals, _options.Margins.Left, y, block);
                        y += trailingHeight;
                        if (blockOffset >= trailingGroup.ContentEndsAt - 0.0001D) blockOffset = trailingGroup.SourceEndsAt;
                    }

                    if (blockOffset < block.Height - 0.0001D) {
                        CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
                        visuals = CreatePageVisuals(pageWidth, pageHeight);
                        y = _options.Margins.Top;
                        currentPageName = block.PageName;
                    }
                }
            }

            if (block.BreakAfter != HtmlPageBreakTarget.None && index < blocks.Count - 1) {
                CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
                EnsurePageSide(block.BreakAfter, pages, ref visuals, ref y, pageWidth, pageHeight, currentPageName);
            }
        }

        CommitPage(pages, visuals, pageWidth, pageHeight, currentPageName);
        return new HtmlRenderDocument(HtmlRenderMode.Paged, ApplyPageMarginContent(pages), _diagnostics, _fonts, _metadata);
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

    private void ApplyBreakBefore(HtmlPageBreakTarget target, ICollection<HtmlRenderPage> pages, ref List<HtmlRenderVisual> visuals, ref double y, double width, double height, string? pageName) {
        if (y > _options.Margins.Top + 0.0001D) {
            CommitPage(pages, visuals, width, height, pageName);
            visuals = CreatePageVisuals(width, height);
            y = _options.Margins.Top;
        }

        EnsurePageSide(target, pages, ref visuals, ref y, width, height, pageName);
    }

    private void EnsurePageSide(HtmlPageBreakTarget target, ICollection<HtmlRenderPage> pages, ref List<HtmlRenderVisual> visuals, ref double y, double width, double height, string? pageName) {
        if (target != HtmlPageBreakTarget.Left && target != HtmlPageBreakTarget.Right) return;
        int nextPageNumber = pages.Count + 1;
        bool nextIsRight = nextPageNumber % 2 != 0;
        bool targetIsRight = target == HtmlPageBreakTarget.Right;
        if (nextIsRight == targetIsRight) return;
        CommitPage(pages, visuals, width, height, pageName);
        visuals = CreatePageVisuals(width, height);
        y = _options.Margins.Top;
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

    private void CommitPage(ICollection<HtmlRenderPage> pages, List<HtmlRenderVisual> visuals, double width, double height, string? pageName) {
        if (pages.Count >= _options.MaxPageCount) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
        }

        bool includeRoot = pages.Count == 0;
        double contentWidth = Math.Max(1D, width - _options.Margins.Left - _options.Margins.Right);
        double contentHeight = Math.Max(1D, height - _options.Margins.Top - _options.Margins.Bottom);
        PrepareGlobalPositionedRequests(includeRoot, width, height, contentWidth, contentHeight);
        AppendGlobalPositionedRequests(visuals, includeRoot, width, height, contentWidth, contentHeight, PositionedPaintBand.Negative);
        AppendGlobalPositionedRequests(visuals, includeRoot, width, height, contentWidth, contentHeight, PositionedPaintBand.NonNegative);
        ApplyViewportOverflow(visuals, width, height);
        pages.Add(new HtmlRenderPage(pages.Count + 1, width, height, visuals, pageName, _fonts));
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

    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available, double? maximumEnd = null) {
        double limit = Math.Min(maximumEnd ?? block.Height, Math.Min(block.Height, start + available));
        double best = start;
        foreach (double offset in block.BreakOffsets) {
            if (offset > start + 0.0001D
                && offset <= limit + 0.0001D
                && IsAllowedLineBreak(block, start, offset)) {
                best = offset;
            }
        }

        return best;
    }

    private static HtmlRenderTrailingGroup? ResolveTrailingGroup(HtmlRenderFlowBlock block, double start, double available, out double fragmentLimit) {
        HtmlRenderTrailingGroup? active = block.TrailingGroups.FirstOrDefault(group => group.AppliesAt(start));
        if (active != null) {
            fragmentLimit = active.ContentEndsAt;
            return active;
        }

        HtmlRenderTrailingGroup? upcoming = block.TrailingGroups
            .Where(group => group.StartsAt > start + 0.0001D && group.StartsAt < start + available - 0.0001D)
            .OrderBy(group => group.StartsAt)
            .FirstOrDefault();
        if (upcoming == null) {
            fragmentLimit = block.Height;
            return null;
        }

        double candidateAvailable = Math.Max(0D, available - upcoming.Height);
        double candidateEnd = FindFragmentEnd(block, start, candidateAvailable, upcoming.ContentEndsAt);
        if (candidateEnd > upcoming.StartsAt + 0.0001D) {
            fragmentLimit = upcoming.ContentEndsAt;
            return upcoming;
        }

        fragmentLimit = upcoming.StartsAt;
        return null;
    }

    private static bool IsAllowedLineBreak(HtmlRenderFlowBlock block, double start, double candidate) {
        foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
            if (!group.Offsets.Any(offset => Math.Abs(offset - candidate) <= 0.0001D)) continue;
            int fragmentLines = group.Offsets.Count(offset => offset > start + 0.0001D && offset <= candidate + 0.0001D);
            int remainingLines = group.Offsets.Count(offset => offset > candidate + 0.0001D);
            if (remainingLines == 0 && candidate < block.Height - 0.0001D) return false;
            return fragmentLines >= group.Orphans && (remainingLines == 0 || remainingLines >= group.Widows);
        }

        return true;
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
                        semanticGroup.LayoutY - start));
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
        double pixelWidth = Math.Ceiling(width * _options.Scale);
        double pixelHeight = Math.Ceiling(height * _options.Scale);
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
        HtmlConversionLossKind lossKind = HtmlConversionLossKind.Approximation) {
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
