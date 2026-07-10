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
    private readonly HtmlRenderResourceSet _resources;
    private readonly HtmlCssPageRuleSet _pageRules;
    private readonly Uri? _baseUri;
    private int _paintOrder;

    internal HtmlRenderLayoutEngine(IHtmlDocument document, IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlRenderResourceSet? resources = null, HtmlCssPageRuleSet? pageRules = null) {
        _document = document;
        _options = options;
        _diagnostics = diagnostics;
        _styleResolver = new HtmlRenderStyleResolver(computedStyles, options);
        _resources = resources ?? new HtmlRenderResourceSet();
        _pageRules = pageRules ?? new HtmlCssPageRuleSet();
        _baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
    }

    internal HtmlRenderDocument Render() {
        IElement root = _document.Body ?? _document.DocumentElement ?? throw new InvalidOperationException("The parsed HTML document has no renderable root element.");
        double surfaceWidth = _options.Mode == HtmlRenderMode.Paged ? _options.PageWidth : _options.ViewportWidth;
        double contentWidth = surfaceWidth - _options.Margins.Left - _options.Margins.Right;
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, contentWidth);
        IReadOnlyList<HtmlRenderFlowBlock> blocks = BuildChildBlocks(root, contentWidth, rootStyle, 0);
        return _options.Mode == HtmlRenderMode.Paged
            ? RenderPaged(blocks)
            : RenderContinuous(blocks);
    }

    private HtmlRenderDocument RenderContinuous(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        double width = _options.ViewportWidth;
        double y = _options.Margins.Top;
        var content = new List<HtmlRenderVisual>();
        foreach (HtmlRenderFlowBlock block in blocks) {
            AddTranslatedVisuals(content, block.Visuals, _options.Margins.Left, y);
            y += block.Height;
        }

        double height = y + _options.Margins.Bottom;
        if (_options.ViewportHeight.HasValue) height = Math.Max(height, _options.ViewportHeight.Value);
        height = Math.Max(1D, height);
        ValidateSurface(width, height);

        var visuals = new List<HtmlRenderVisual> {
            CreatePageBackground(width, height)
        };
        visuals.AddRange(content);
        var page = new HtmlRenderPage(1, width, height, visuals);
        return new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, _diagnostics);
    }

    private HtmlRenderDocument RenderPaged(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        double pageWidth = _options.PageWidth;
        double pageHeight = _options.PageHeight;
        double contentHeight = pageHeight - _options.Margins.Top - _options.Margins.Bottom;
        ValidateSurface(pageWidth, pageHeight);

        var pages = new List<HtmlRenderPage>();
        var visuals = CreatePageVisuals(pageWidth, pageHeight);
        double y = _options.Margins.Top;
        for (int index = 0; index < blocks.Count; index++) {
            HtmlRenderFlowBlock block = blocks[index];
            bool hasPageContent = visuals.Count > 1;
            if (block.BreakBefore && hasPageContent) {
                CommitPage(pages, visuals, pageWidth, pageHeight);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
                hasPageContent = false;
            }

            if (block.Height <= contentHeight && hasPageContent && y + block.Height > pageHeight - _options.Margins.Bottom) {
                CommitPage(pages, visuals, pageWidth, pageHeight);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
            }

            if (block.Height <= pageHeight - _options.Margins.Bottom - y) {
                AddTranslatedVisuals(visuals, block.Visuals, _options.Margins.Left, y);
                y += block.Height;
            } else {
                double blockOffset = 0D;
                while (blockOffset < block.Height - 0.0001D) {
                    HtmlRenderContinuationGroup? continuationGroup = block.ContinuationGroups.FirstOrDefault(group => group.AppliesAt(blockOffset));
                    bool repeatContinuation = blockOffset > 0.0001D && continuationGroup != null && continuationGroup.Visuals.Count > 0 && continuationGroup.Height > 0D;
                    double continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                    double available = pageHeight - _options.Margins.Bottom - y - continuationHeight;
                    if (available <= 0.0001D) {
                        if (visuals.Count > 1) {
                            CommitPage(pages, visuals, pageWidth, pageHeight);
                            visuals = CreatePageVisuals(pageWidth, pageHeight);
                            y = _options.Margins.Top;
                            continue;
                        }

                        repeatContinuation = false;
                        continuationHeight = 0D;
                        available = contentHeight;
                        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "A repeated table header was suppressed because it consumed the available page content area.", HtmlDiagnosticSeverity.Warning, block.Source);
                    }

                    double fragmentEnd = FindFragmentEnd(block, blockOffset, available);
                    if (fragmentEnd <= blockOffset + 0.0001D) {
                        if (visuals.Count > 1) {
                            CommitPage(pages, visuals, pageWidth, pageHeight);
                            visuals = CreatePageVisuals(pageWidth, pageHeight);
                            y = _options.Margins.Top;
                            continue;
                        }

                        if (repeatContinuation) {
                            double fallbackAvailable = pageHeight - _options.Margins.Bottom - y;
                            double fallbackEnd = FindFragmentEnd(block, blockOffset, fallbackAvailable);
                            if (fallbackEnd > blockOffset + 0.0001D) {
                                repeatContinuation = false;
                                continuationHeight = 0D;
                                available = fallbackAvailable;
                                fragmentEnd = fallbackEnd;
                                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "A repeated table header was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
                            }
                        }

                        if (fragmentEnd <= blockOffset + 0.0001D) {
                            fragmentEnd = Math.Min(block.Height, blockOffset + available);
                            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ForcedFragment, "A layout block had no safe break opportunity within one page and was force-fragmented.", HtmlDiagnosticSeverity.Warning, block.Source);
                        }
                    }

                    if (repeatContinuation) {
                        AddTranslatedVisuals(visuals, continuationGroup!.Visuals, _options.Margins.Left, y);
                        y += continuationHeight;
                    }

                    IReadOnlyList<HtmlRenderVisual> fragment = SliceBlockVisuals(block, blockOffset, fragmentEnd);
                    AddTranslatedVisuals(visuals, fragment, _options.Margins.Left, y);
                    y += fragmentEnd - blockOffset;
                    blockOffset = fragmentEnd;
                    if (blockOffset < block.Height - 0.0001D) {
                        CommitPage(pages, visuals, pageWidth, pageHeight);
                        visuals = CreatePageVisuals(pageWidth, pageHeight);
                        y = _options.Margins.Top;
                    }
                }
            }

            if (block.BreakAfter && index < blocks.Count - 1) {
                CommitPage(pages, visuals, pageWidth, pageHeight);
                visuals = CreatePageVisuals(pageWidth, pageHeight);
                y = _options.Margins.Top;
            }
        }

        CommitPage(pages, visuals, pageWidth, pageHeight);
        return new HtmlRenderDocument(HtmlRenderMode.Paged, ApplyPageMarginContent(pages), _diagnostics);
    }

    private List<HtmlRenderVisual> CreatePageVisuals(double width, double height) => new List<HtmlRenderVisual> { CreatePageBackground(width, height) };

    private HtmlRenderShape CreatePageBackground(double width, double height) {
        OfficeShape background = OfficeShape.Rectangle(width, height);
        background.FillColor = _options.BackgroundColor;
        background.StrokeWidth = 0D;
        return new HtmlRenderShape(background, 0D, 0D, _paintOrder++, source: "render-surface");
    }

    private void CommitPage(ICollection<HtmlRenderPage> pages, List<HtmlRenderVisual> visuals, double width, double height) {
        if (pages.Count >= _options.MaxPageCount) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
        }

        pages.Add(new HtmlRenderPage(pages.Count + 1, width, height, visuals));
    }

    private void AddTranslatedVisuals(ICollection<HtmlRenderVisual> target, IEnumerable<HtmlRenderVisual> source, double offsetX, double offsetY) {
        foreach (HtmlRenderVisual visual in source) {
            target.Add(visual.Translate(offsetX, offsetY, _paintOrder++));
        }
    }

    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available) {
        double limit = Math.Min(block.Height, start + available);
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
        var fragment = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in block.Visuals) {
            double visualTop = visual.Y;
            double visualBottom = visual.Y + visual.Height;
            double intersectionTop = Math.Max(start, visualTop);
            double intersectionBottom = Math.Min(end, visualBottom);
            if (intersectionBottom <= intersectionTop + 0.0001D) continue;

            bool fullyContained = visualTop >= start - 0.0001D && visualBottom <= end + 0.0001D;
            if (fullyContained) {
                fragment.Add(visual.Translate(0D, -start, fragment.Count));
                continue;
            }

            if (visual is HtmlRenderShape shape
                && (shape.Shape.Kind == OfficeShapeKind.Rectangle || shape.Shape.Kind == OfficeShapeKind.RoundedRectangle)) {
                OfficeShape sliced = shape.Shape.Clone();
                sliced.Height = intersectionBottom - intersectionTop;
                if (sliced.Kind == OfficeShapeKind.RoundedRectangle) sliced.CornerRadius = Math.Min(sliced.CornerRadius, sliced.Height / 2D);
                fragment.Add(new HtmlRenderShape(sliced, shape.X, intersectionTop - start, fragment.Count, shape.LinkUri, shape.Source));
                continue;
            }

            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "A visual crossing a forced page boundary could not be represented safely in the current fragment.", HtmlDiagnosticSeverity.Warning, visual.Source, visual.Kind.ToString());
        }

        return fragment;
    }

    private void ValidateSurface(double width, double height) {
        long pixelWidth = (long)Math.Ceiling(width * _options.Scale);
        long pixelHeight = (long)Math.Ceiling(height * _options.Scale);
        if (pixelWidth > _options.MaxSurfaceWidth || pixelHeight > _options.MaxSurfaceHeight) {
            throw new InvalidOperationException("HTML rendering exceeded the configured maximum image surface dimensions.");
        }
    }

    private void AddUnsupported(string code, string message, IElement element, string? detail = null) {
        _diagnostics.Add(ComponentName, code, message, HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element), detail);
    }
}
