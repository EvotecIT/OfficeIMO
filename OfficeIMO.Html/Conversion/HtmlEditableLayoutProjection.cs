using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Formatting regions a native destination can request from the shared projector.</summary>
[Flags]
public enum HtmlEditableLayoutRegionKinds {
    /// <summary>No editable regions.</summary>
    None = 0,
    /// <summary>Absolutely or fixed-positioned boxes.</summary>
    Positioned = 1,
    /// <summary>Left and right floats.</summary>
    Floating = 2,
    /// <summary>Flex formatting contexts.</summary>
    Flex = 4,
    /// <summary>Grid formatting contexts.</summary>
    Grid = 8,
    /// <summary>All supported bounded formatting contexts.</summary>
    All = Positioned | Floating | Flex | Grid
}

/// <summary>Stable diagnostics for editable native HTML layout projection.</summary>
public static class HtmlEditableLayoutDiagnosticCodes {
    /// <summary>A bounded region was retained for native destination projection.</summary>
    public const string RegionProjected = "HtmlEditableLayoutRegionProjected";
    /// <summary>A region crossed render surfaces and stayed in semantic flow rather than becoming ambiguous native geometry.</summary>
    public const string RegionFragmented = "HtmlEditableLayoutRegionFragmented";
    /// <summary>A destination flattened extra background layers.</summary>
    public const string BackgroundLayersFlattened = "HtmlEditableLayoutBackgroundLayersFlattened";
    /// <summary>A destination omitted an unsupported native effect while retaining editable content.</summary>
    public const string EffectUnsupported = "HtmlEditableLayoutEffectUnsupported";
    /// <summary>A native destination moved or resized a region to preserve existing editable content.</summary>
    public const string PlacementSimplified = "HtmlEditableLayoutPlacementSimplified";
    /// <summary>A region image could not be retained by the destination's native editable representation.</summary>
    public const string RegionImageOmitted = "HtmlEditableLayoutRegionImageOmitted";
}

/// <summary>Shared rendered placement plan consumed by thin native target adapters.</summary>
public sealed class HtmlEditableLayoutProjection {
    internal HtmlEditableLayoutProjection(
        IHtmlDocument remainingDocument,
        HtmlRenderDocument renderedDocument,
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> sourceImages,
        IReadOnlyDictionary<string, IHtmlImageElement> sourceImagesByRenderKey,
        IReadOnlyList<HtmlDiagnostic> diagnostics) {
        RemainingDocument = remainingDocument;
        RenderedDocument = renderedDocument;
        Regions = regions;
        _sourceImages = sourceImages;
        _sourceImagesByRenderKey = sourceImagesByRenderKey;
        Diagnostics = diagnostics;
    }

    private readonly IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> _sourceImages;
    private readonly IReadOnlyDictionary<string, IHtmlImageElement> _sourceImagesByRenderKey;

    /// <summary>Backend-neutral rendered evidence used to derive native geometry.</summary>
    public HtmlRenderDocument RenderedDocument { get; }
    /// <summary>Bounded, single-surface editable regions in source order.</summary>
    public IReadOnlyList<HtmlRenderLayoutRegion> Regions { get; }
    /// <summary>Projection diagnostics, including stable fragmentation decisions.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics { get; }
    internal IHtmlDocument RemainingDocument { get; }
    internal IReadOnlyList<IHtmlImageElement> GetSourceImages(HtmlRenderLayoutRegion region) =>
        _sourceImages.TryGetValue(region.SourceKey, out IReadOnlyList<IHtmlImageElement>? images)
            ? images
            : Array.Empty<IHtmlImageElement>();
    internal IHtmlImageElement? GetSourceImage(HtmlRenderImage renderedImage) =>
        renderedImage.Source != null
            && _sourceImagesByRenderKey.TryGetValue(renderedImage.Source, out IHtmlImageElement? sourceImage)
                ? sourceImage
                : null;
}

/// <summary>Creates one shared editable-layout plan for DOCX, RTF, XLSX, and PPTX adapters.</summary>
public static partial class HtmlEditableLayoutProjector {
    internal const string RegionAttribute = "data-officeimo-editable-layout-region";
    internal const string ImageAttribute = "data-officeimo-editable-layout-image";
    private const string ImageSourcePrefix = "img[officeimo-layout-image=";
    private static readonly HashSet<string> SemanticRichElementNames = new(StringComparer.OrdinalIgnoreCase) {
        "a", "abbr", "audio", "b", "blockquote", "br", "button", "canvas", "cite", "code", "dd", "del",
        "details", "dfn", "dl", "dt", "em", "embed", "fieldset", "figure", "figcaption", "form", "h1", "h2",
        "h3", "h4", "h5", "h6", "hr", "i", "iframe", "input", "ins", "kbd", "label", "li", "mark",
        "meter", "object", "ol", "p", "picture", "pre", "progress", "q", "s", "samp", "select", "strong", "sub", "summary",
        "sup", "svg", "table", "textarea", "time", "u", "ul", "var", "video",
        "math", "ruby", "rb", "rp", "rt", "rtc"
    };
    private static readonly string[] RichTextStyleProperties = {
        "color", "direction", "font-family", "font-size", "font-style", "font-variant", "font-weight",
        "letter-spacing", "line-height",
        "text-decoration", "text-decoration-color", "text-decoration-line", "text-decoration-style", "text-shadow",
        "text-align", "text-indent", "text-transform", "unicode-bidi", "vertical-align", "white-space", "word-spacing"
    };
    private static readonly string[] RichDescendantVisualStyleProperties = {
        "background-color", "border-bottom-color", "border-bottom-style", "border-bottom-width",
        "border-left-color", "border-left-style", "border-left-width", "border-right-color", "border-right-style",
        "border-right-width", "border-top-color", "border-top-style", "border-top-width"
    };

    /// <summary>Projects bounded positioned, floating, flex, and grid regions through the managed layout engine.</summary>
    public static HtmlEditableLayoutProjection Project(
        HtmlConversionDocument document,
        HtmlRenderOptions? renderOptions = null,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen,
        HtmlEditableLayoutRegionKinds regionKinds = HtmlEditableLayoutRegionKinds.All,
        int? maximumEditableSurfaceNumber = null) => ProjectCore(
            document, renderOptions, mediaContext, regionKinds, maximumEditableSurfaceNumber,
             maximumEditableContinuousSurfaceHeight: null,
             preserveMixedInlineContent: false,
             preserveNestedImagePlacement: false,
             preserveMixedInlineEdgeSequences: false,
             preserveRegionsBeforeForcedPageBreaks: false,
             preserveRegionsAfterForcedPageBreaks: false,
             preserveUnrenderedSourceImages: false,
             limits: null);

    internal static HtmlEditableLayoutProjection ProjectPreservingMixedInlineContent(
        HtmlConversionDocument document,
        HtmlRenderOptions? renderOptions = null,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen,
        HtmlEditableLayoutRegionKinds regionKinds = HtmlEditableLayoutRegionKinds.All,
        int? maximumEditableSurfaceNumber = null,
         double? maximumEditableContinuousSurfaceHeight = null,
         HtmlConversionLimits? limits = null,
         bool preserveNestedImagePlacement = true,
         bool preserveMixedInlineEdgeSequences = true,
         bool preserveRegionsBeforeForcedPageBreaks = false,
         bool preserveRegionsAfterForcedPageBreaks = false,
         bool preserveUnrenderedSourceImages = false) => ProjectCore(
             document, renderOptions, mediaContext, regionKinds, maximumEditableSurfaceNumber,
             maximumEditableContinuousSurfaceHeight,
             preserveMixedInlineContent: true,
             preserveNestedImagePlacement,
             preserveMixedInlineEdgeSequences,
             preserveRegionsBeforeForcedPageBreaks,
             preserveRegionsAfterForcedPageBreaks,
             preserveUnrenderedSourceImages,
             limits);

    private static HtmlEditableLayoutProjection ProjectCore(
        HtmlConversionDocument document,
        HtmlRenderOptions? renderOptions,
        HtmlCssMediaContext mediaContext,
        HtmlEditableLayoutRegionKinds regionKinds,
        int? maximumEditableSurfaceNumber,
        double? maximumEditableContinuousSurfaceHeight,
         bool preserveMixedInlineContent,
         bool preserveNestedImagePlacement,
         bool preserveMixedInlineEdgeSequences,
         bool preserveRegionsBeforeForcedPageBreaks,
         bool preserveRegionsAfterForcedPageBreaks,
         bool preserveUnrenderedSourceImages,
         HtmlConversionLimits? limits) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if ((regionKinds & ~HtmlEditableLayoutRegionKinds.All) != 0) throw new ArgumentOutOfRangeException(nameof(regionKinds));
        if (maximumEditableSurfaceNumber < 0) throw new ArgumentOutOfRangeException(nameof(maximumEditableSurfaceNumber));
        if (maximumEditableContinuousSurfaceHeight <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(maximumEditableContinuousSurfaceHeight));
        }
        IHtmlDocument adapterDocument = document.CreateSourceDocumentForConversion();
        IReadOnlyList<HtmlGenericSectionProjection> semanticSections =
            HtmlGenericDocumentProjector.CreateSections(adapterDocument);
        IReadOnlyList<IElement> semanticTables = HtmlGenericDocumentProjector.SelectRootTables(adapterDocument);
        HtmlRenderOptions options = renderOptions?.Clone() ?? new HtmlRenderOptions();
        HtmlConversionLimits effectiveLimits = limits == null
            ? document.Limits.Clone()
            : HtmlConversionLimits.Intersect(document.Limits, limits);
        options.Mode = mediaContext == HtmlCssMediaContext.Print
            ? HtmlRenderMode.Paged
            : HtmlRenderMode.Continuous;
        IReadOnlyList<IHtmlStyleElement> callerStyles = HtmlRenderAdditionalStylesheetApplier.Apply(
            adapterDocument, options.AdditionalStylesheets.ToList());
        options.AdditionalStylesheets.Clear();
        HtmlComputedStyleSet computedStyles = HtmlComputedStyleEngine.ComputeForRendering(
            adapterDocument, options, effectiveLimits);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = computedStyles.Elements;
        IReadOnlyList<IElement> elementsInDocumentOrder = adapterDocument.QuerySelectorAll("*").ToList();
        IReadOnlyDictionary<IElement, int> documentOrder = elementsInDocumentOrder
            .Select((element, index) => new { Element = element, Index = index })
            .ToDictionary(item => item.Element, item => item.Index);
        IReadOnlyList<double> forcedPageBreakBoundaries = CreateForcedPageBreakBoundaries(
            elementsInDocumentOrder, styles);
        int key = 0;
        var candidateElements = new Dictionary<string, IElement>(StringComparer.Ordinal);
        var semanticFlowRoots = new HashSet<IElement>();
        var semanticRichCandidates = new List<IElement>();
        var multipleBlockContentCandidates = new List<IElement>();
        var inheritedTypographyCandidates = new List<IElement>();
        var multiChildLayoutKeys = new HashSet<string>(StringComparer.Ordinal);
        var nestedLayoutPlacementKeys = new HashSet<string>(StringComparer.Ordinal);
        var bookmarkTargetCandidates = new List<IElement>();
        var commentBearingCandidates = new List<IElement>();
        var paintHiddenCandidates = new List<IElement>();
        var generatedContentCandidates = new List<IElement>();
        var forcedPageBreakCandidates = new List<(IElement Element, string Detail)>();
        var effectCandidates = new List<(IElement Element, string Detail)>();
        var mixedInlineImageCandidates = new List<IElement>();
        foreach (IElement element in elementsInDocumentOrder) {
            if (!styles.TryGetValue(element, out HtmlComputedStyle? style)
                || !TryGetCandidateKind(element, style, out HtmlEditableLayoutRegionKinds candidateKind)
                || (regionKinds & candidateKind) == 0) continue;
            if (HasSemanticFlowAncestor(element, semanticFlowRoots)) continue;
            if (!IsProjectionElementVisible(element, element, styles)) {
                paintHiddenCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (TryGetForcedPageBreakOwnershipDetail(
                    documentOrder[element],
                    forcedPageBreakBoundaries,
                    preserveRegionsBeforeForcedPageBreaks,
                    preserveRegionsAfterForcedPageBreaks,
                    out string forcedBreakDetail)) {
                forcedPageBreakCandidates.Add((element, forcedBreakDetail));
                semanticFlowRoots.Add(element);
                continue;
            }
            if (ContainsGeneratedPseudoContent(element, computedStyles)) {
                generatedContentCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (ContainsHtmlComment(element)) {
                commentBearingCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (ContainsBookmarkTarget(element)) {
                bookmarkTargetCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (ContainsMultipleVisibleBlockContentItems(element, styles)) {
                multipleBlockContentCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (TryGetNonNativeRegionEffect(element, styles, out string effectDetail)) {
                effectCandidates.Add((element, effectDetail));
                semanticFlowRoots.Add(element);
                continue;
            }
            if (preserveMixedInlineContent
                && ContainsMixedInlineImageContent(element, styles, preserveMixedInlineEdgeSequences)) {
                mixedInlineImageCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (HasInheritedRichTextStyle(element, styles)) {
                inheritedTypographyCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            if (ContainsSemanticRichContent(element, styles)) {
                semanticRichCandidates.Add(element);
                semanticFlowRoots.Add(element);
                continue;
            }
            string sourceKey = (++key).ToString(System.Globalization.CultureInfo.InvariantCulture);
            SetRegionSourceKey(element, sourceKey);
            candidateElements[sourceKey] = element;
            if (IsFlexOrGridDisplay(style)
                && HasMultipleVisibleLayoutChildren(element, styles)) {
                multiChildLayoutKeys.Add(sourceKey);
            }
            if (ContainsNestedLayoutPlacement(element, styles, includeImages: preserveNestedImagePlacement)) {
                nestedLayoutPlacementKeys.Add(sourceKey);
            }
        }
        int imageKey = 0;
        foreach (IElement image in adapterDocument.QuerySelectorAll("img")) {
            SetImageSourceKey(image, (++imageKey).ToString(
                System.Globalization.CultureInfo.InvariantCulture));
        }

        options.EnableEditableLayoutRegions = true;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            adapterDocument, options, document, effectiveLimits);
        double continuousSurfaceHeight = rendered.Pages.Count == 0
            ? 0D
            : rendered.Pages.Max(page => page.Height);
        bool continuousPageOwnershipUnavailable = options.Mode == HtmlRenderMode.Continuous
            && maximumEditableContinuousSurfaceHeight.HasValue
            && continuousSurfaceHeight > maximumEditableContinuousSurfaceHeight.Value;
        var occurrences = new List<EditableLayoutRegionOccurrence>();
        foreach (HtmlRenderPage page in rendered.Pages) {
            occurrences.AddRange(EnumerateRegions(page.Scene, page.PageNumber, null));
        }

        var diagnostics = new HtmlDiagnosticReport();
        foreach (IElement element in commentBearingCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so its raw HTML comments could remain available to the destination.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "htmlComment=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in bookmarkTargetCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so its bookmark target would remain addressable.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "bookmarkTarget=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in paintHiddenCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "A paint-hidden editable layout region stayed in semantic flow rather than becoming visible destination-native geometry.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "paintVisible=false; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in generatedContentCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so generated pseudo-element content and paint would remain intact.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                 "generatedContent=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach ((IElement element, string detail) in forcedPageBreakCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow because its destination-native anchor cannot cross a forced page break safely.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                detail + "; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in semanticRichCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so rich document content would not be flattened.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "semanticContent=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in multipleBlockContentCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so visible block boundaries would not be flattened into one native text run.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "multipleBlockChildren=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in inheritedTypographyCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so inherited typography would remain intact.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "inheritedTypography=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach ((IElement element, string detail) in effectCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                "An editable layout region stayed in semantic flow because its layout, box model, or paint effect has no exact destination-native representation.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                detail + "; semanticFlow=true", OfficeConversionLossKind.Approximation);
        }
        foreach (IElement element in mixedInlineImageCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so inline picture and text order would remain intact.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "mixedInlinePictures=true", OfficeConversionLossKind.Approximation);
        }
        var accepted = new List<HtmlRenderLayoutRegion>();
        foreach (IGrouping<string, EditableLayoutRegionOccurrence> group in occurrences.GroupBy(item => item.Region.SourceKey)) {
            int occurrenceCount = group.Count();
            int pageCount = group.Select(item => item.Page).Distinct().Count();
            if (occurrenceCount != 1) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.RegionFragmented,
                    "An editable layout region produced multiple rendered fragments and remained in semantic flow.",
                    HtmlDiagnosticSeverity.Warning, group.First().Region.Source,
                    "occurrences=" + occurrenceCount + "; surfaces=" + pageCount, OfficeConversionLossKind.Approximation);
                continue;
            }
            HtmlRenderLayoutRegion selected = group
                .OrderBy(item => RegionPriority(item.Region.RegionKind))
                .ThenByDescending(item => item.Region.Width * item.Region.Height)
                .Select(item => item.Region)
                .First();
            selected.SurfaceNumber = group.First().Page;
            EditableLayoutRegionOccurrence selectedOccurrence = group.First(item => ReferenceEquals(item.Region, selected));
            selected.SemanticSectionOriginX = selectedOccurrence.SectionOriginX;
            selected.SemanticSectionOriginY = selectedOccurrence.SectionOriginY;
            selected.SemanticTableOriginX = selectedOccurrence.TableOriginX;
            selected.SemanticTableOriginY = selectedOccurrence.TableOriginY;
            if (preserveUnrenderedSourceImages
                && ContainsUnrenderedSourceImage(selected, candidateElements[selected.SourceKey], styles)) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "An editable layout region stayed in semantic flow because one of its source pictures lacked rendered occurrence metadata.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "unrenderedRegionImage=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
                continue;
            }
            bool selectedPageOwnershipUnavailable = continuousPageOwnershipUnavailable
                || (options.Mode == HtmlRenderMode.Continuous
                    && maximumEditableContinuousSurfaceHeight.HasValue
                    && selected.Y >= maximumEditableContinuousSurfaceHeight.Value);
            if (selectedPageOwnershipUnavailable) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "An editable layout region stayed in semantic flow because a multi-page continuous coordinate cannot be mapped to one page-relative native anchor.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "continuousSurfaceHeight=" + continuousSurfaceHeight.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; regionY=" + selected.Y.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        + "; maximumPageHeight=" + maximumEditableContinuousSurfaceHeight!.Value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
                    OfficeConversionLossKind.Approximation);
                continue;
            }
            if (maximumEditableSurfaceNumber.HasValue
                && selected.SurfaceNumber > maximumEditableSurfaceNumber.Value) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "An editable layout region stayed in semantic flow because its rendered surface cannot be mapped natively by this destination.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "surface=" + selected.SurfaceNumber + "; maximumEditableSurface=" + maximumEditableSurfaceNumber.Value,
                    OfficeConversionLossKind.Approximation);
                continue;
            }
            if (nestedLayoutPlacementKeys.Contains(selected.SourceKey)) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "An editable layout region stayed in semantic flow because flattening its nested layout placement would lose child geometry.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "nestedLayoutPlacement=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
                continue;
            }
            if (HasAcceptedAncestor(candidateElements[selected.SourceKey], nestedLayoutPlacementKeys)) continue;
            if (multiChildLayoutKeys.Contains(selected.SourceKey)) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "A flex or grid editable layout region stayed in semantic flow because its child placement cannot be represented by one native text container.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "multipleLayoutChildren=true; semanticFlow=true", OfficeConversionLossKind.Approximation);
                continue;
            }
            if (HasAcceptedAncestor(candidateElements[selected.SourceKey], multiChildLayoutKeys)) continue;
            accepted.Add(selected);
        }

        var sectionOwned = new List<HtmlRenderLayoutRegion>();
        foreach (HtmlRenderLayoutRegion region in accepted) {
            IElement element = candidateElements[region.SourceKey];
            if (!TryGetSemanticSectionNumber(element, semanticSections, out int sectionNumber, out int matchCount)) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.RegionFragmented,
                    "An editable layout region crossed generic semantic section boundaries and remained in semantic flow.",
                    HtmlDiagnosticSeverity.Warning, region.Source,
                    "semanticSections=" + matchCount, OfficeConversionLossKind.Approximation);
                continue;
            }
            region.SemanticSectionNumber = sectionNumber;
            if (semanticSections.Count <= 1) {
                region.SemanticSectionOriginX = 0D;
                region.SemanticSectionOriginY = 0D;
            }
            region.SemanticTableNumber = TryGetOwningElementNumber(
                candidateElements[region.SourceKey], semanticTables);
            sectionOwned.Add(region);
        }
        accepted = sectionOwned;

        var preliminaryKeys = new HashSet<string>(accepted.Select(region => region.SourceKey), StringComparer.Ordinal);
        accepted = accepted.Where(region => !HasAcceptedAncestor(
                candidateElements[region.SourceKey], preliminaryKeys))
            .ToList();
        foreach (HtmlRenderLayoutRegion region in accepted) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.RegionProjected,
                "A bounded HTML layout region is available for native editable projection.",
                HtmlDiagnosticSeverity.Info, region.Source, region.RegionKind.ToString());
        }

        var acceptedKeys = new HashSet<string>(accepted.Select(region => region.SourceKey), StringComparer.Ordinal);
        var semanticSectionOwnerKeys = new HashSet<string>(accepted
            .Where(region => IsSemanticSectionOwner(candidateElements[region.SourceKey], adapterDocument))
            .Select(region => region.SourceKey), StringComparer.Ordinal);
        IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> orderedSourceImages = accepted.ToDictionary(
           region => region.SourceKey,
            region => CreateOrderedSourceImages(region, candidateElements[region.SourceKey], styles),
            StringComparer.Ordinal);
        var originalRegionAttributeValues = candidateElements.ToDictionary(
            pair => pair.Key,
            pair => pair.Value.GetAttribute(RegionAttribute),
            StringComparer.Ordinal);
        foreach (KeyValuePair<string, IElement> pair in candidateElements) {
            pair.Value.SetAttribute(RegionAttribute, pair.Key);
        }
        IReadOnlyList<IHtmlImageElement> markedImages = adapterDocument.QuerySelectorAll("img")
            .OfType<IHtmlImageElement>()
            .Where(image => GetImageSourceKey(image) != null)
            .ToList();
        var originalImageAttributeValues = markedImages.ToDictionary(
            image => GetImageSourceKey(image)!,
            image => image.GetAttribute(ImageAttribute),
            StringComparer.Ordinal);
        foreach (IHtmlImageElement image in markedImages) {
            image.SetAttribute(ImageAttribute, GetImageSourceKey(image)!);
        }
        foreach (IHtmlStyleElement style in callerStyles) style.Remove();
        IHtmlDocument remainingDocument = document.CreatePolicyNormalizedDocumentForConversion(adapterDocument);
        IReadOnlyDictionary<string, IHtmlImageElement> normalizedImagesByMarker = remainingDocument
            .QuerySelectorAll("img[" + ImageAttribute + "]")
            .OfType<IHtmlImageElement>()
            .Where(image => !string.IsNullOrWhiteSpace(image.GetAttribute(ImageAttribute)))
            .ToDictionary(image => image.GetAttribute(ImageAttribute)!, image => image, StringComparer.Ordinal);
        IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> sourceImages = orderedSourceImages.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<IHtmlImageElement>)pair.Value
                .Select(image => image.GetAttribute(ImageAttribute))
                .Where(marker => marker != null && normalizedImagesByMarker.ContainsKey(marker))
                .Select(marker => normalizedImagesByMarker[marker!])
                .ToArray(),
            StringComparer.Ordinal);
        IReadOnlyDictionary<string, IHtmlImageElement> sourceImagesByRenderKey = sourceImages.Values
           .SelectMany(images => images)
            .ToDictionary(image => DescribeImageSource(image.GetAttribute(ImageAttribute)), image => image,
                StringComparer.Ordinal);
        foreach (KeyValuePair<string, IHtmlImageElement> pair in normalizedImagesByMarker) {
            RestoreAuthoredAttribute(pair.Value, ImageAttribute, pair.Key, originalImageAttributeValues);
        }
        foreach (IHtmlImageElement image in markedImages) {
            RestoreAuthoredAttribute(image, ImageAttribute, GetImageSourceKey(image), originalImageAttributeValues);
        }
        foreach (KeyValuePair<string, IElement> pair in candidateElements) {
            RestoreAuthoredAttribute(pair.Value, RegionAttribute, pair.Key, originalRegionAttributeValues);
        }
        foreach (IElement element in remainingDocument.QuerySelectorAll("[" + RegionAttribute + "]").ToArray()) {
            string? sourceKey = element.GetAttribute(RegionAttribute);
            if (sourceKey != null && acceptedKeys.Contains(sourceKey)) {
                if (semanticSectionOwnerKeys.Contains(sourceKey)) {
                    foreach (INode child in element.ChildNodes.ToArray()) element.RemoveChild(child);
                    RestoreAuthoredAttribute(element, RegionAttribute, sourceKey, originalRegionAttributeValues);
                } else {
                    element.Remove();
                }
            } else {
                RestoreAuthoredAttribute(element, RegionAttribute, sourceKey, originalRegionAttributeValues);
            }
        }
        IReadOnlyList<HtmlDiagnostic> projectionDiagnostics = rendered.Diagnostics
            .Where(renderDiagnostic => !document.Diagnostics.Any(sourceDiagnostic =>
                DiagnosticsAreEquivalent(sourceDiagnostic, renderDiagnostic)))
            .Concat(diagnostics.Diagnostics)
            .ToArray();
        return new HtmlEditableLayoutProjection(remainingDocument, rendered, accepted.AsReadOnly(), sourceImages,
            sourceImagesByRenderKey, projectionDiagnostics);
    }

    internal static HtmlSemanticDocument BuildRemainingSemanticDocument(
        HtmlEditableLayoutProjection projection,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits limits) =>
        HtmlSemanticDocumentBuilder.FromDocument(projection.RemainingDocument, mediaContext, limits);

    internal static bool MayContainEditableLayoutRegions(
        HtmlConversionDocument document,
        HtmlEditableLayoutRegionKinds regionKinds,
        IEnumerable<string>? additionalStylesheets = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (MayContainEditableLayoutRegions(document.SourceHtml, regionKinds)) return true;
        return additionalStylesheets != null
            && additionalStylesheets.Any(stylesheet =>
                MayContainEditableLayoutRegions(stylesheet, regionKinds));
    }

    private static bool MayContainEditableLayoutRegions(
        string? html,
        HtmlEditableLayoutRegionKinds regionKinds) {
        if (string.IsNullOrWhiteSpace(html)) return false;
        string content = html!;
        return ((regionKinds & HtmlEditableLayoutRegionKinds.Positioned) != 0
                && content.IndexOf("position", StringComparison.OrdinalIgnoreCase) >= 0)
            || ((regionKinds & HtmlEditableLayoutRegionKinds.Floating) != 0
                && content.IndexOf("float", StringComparison.OrdinalIgnoreCase) >= 0)
            || ((regionKinds & (HtmlEditableLayoutRegionKinds.Flex | HtmlEditableLayoutRegionKinds.Grid)) != 0
                && content.IndexOf("display", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    internal static IEnumerable<(HtmlRenderImage Image, double Opacity)> EnumerateImages(
        IEnumerable<HtmlRenderVisual> visuals,
        bool includeBackgroundImages,
        double opacity = 1D) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderImage image
                && (includeBackgroundImages || !IsBackgroundImage(image))) {
                yield return (image, opacity);
            }
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                _ => null
            };
            if (children == null) continue;
            double childOpacity = visual is HtmlRenderEffectGroup group ? opacity * group.Opacity : opacity;
            foreach ((HtmlRenderImage Image, double Opacity) child in EnumerateImages(
                         children, includeBackgroundImages, childOpacity)) {
                yield return child;
            }
        }
    }

    private static bool TryGetCandidateKind(
        IElement element,
        HtmlComputedStyle style,
        out HtmlEditableLayoutRegionKinds kind) {
        switch (element.LocalName.ToLowerInvariant()) {
            case "div":
            case "section":
            case "article":
            case "aside":
            case "header":
            case "footer":
            case "nav":
            case "main":
            case "figure":
            case "figcaption":
                break;
            default:
                kind = HtmlEditableLayoutRegionKinds.None;
                return false;
        }
        string position = style.GetValue("position").Trim().ToLowerInvariant();
        string floatSide = style.GetValue("float").Trim().ToLowerInvariant();
        string display = style.GetValue("display").Trim().ToLowerInvariant();
        if (position == "absolute" || position == "fixed") kind = HtmlEditableLayoutRegionKinds.Positioned;
        else if (floatSide == "left" || floatSide == "right") kind = HtmlEditableLayoutRegionKinds.Floating;
        else if (display == "flex" || display == "inline-flex") kind = HtmlEditableLayoutRegionKinds.Flex;
        else if (display == "grid" || display == "inline-grid") kind = HtmlEditableLayoutRegionKinds.Grid;
        else {
            kind = HtmlEditableLayoutRegionKinds.None;
            return false;
        }
        return true;
    }

    private static int RegionPriority(HtmlRenderLayoutRegionKind kind) => kind switch {
        HtmlRenderLayoutRegionKind.Positioned => 0,
        HtmlRenderLayoutRegionKind.Floating => 1,
        HtmlRenderLayoutRegionKind.Grid => 2,
        _ => 3
    };

    private static bool HasAcceptedAncestor(IElement element, ISet<string> acceptedKeys) {
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            string? sourceKey = GetRegionSourceKey(ancestor);
            if (sourceKey != null && acceptedKeys.Contains(sourceKey)) return true;
        }
        return false;
    }

    private static bool TryGetSemanticSectionNumber(
        IElement element,
        IReadOnlyList<HtmlGenericSectionProjection> sections,
        out int sectionNumber,
        out int matchCount) {
        var matches = new List<int>();
        for (int index = 0; index < sections.Count; index++) {
            if (sections[index].Blocks.Any(block => ReferenceEquals(block, element)
                    || block.Contains(element)
                    || element.Contains(block))) {
                matches.Add(index + 1);
            }
        }
        if (matches.Count == 0 && sections.Count == 1) matches.Add(1);
        matchCount = matches.Count;
        sectionNumber = matches.Count == 1 ? matches[0] : 0;
        return matches.Count == 1;
    }

    internal static bool IsBackgroundImage(HtmlRenderImage image) =>
        image.Source?.IndexOf(":background-image", StringComparison.Ordinal) >= 0;

    internal static string DescribeImageSource(string? key) =>
        ImageSourcePrefix + (key ?? string.Empty) + "]";

    private static IReadOnlyList<IHtmlImageElement> CreateOrderedSourceImages(
        HtmlRenderLayoutRegion region,
        IElement regionElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        var sourceItems = regionElement.QuerySelectorAll("img")
            .OfType<IHtmlImageElement>()
            .Where(image => IsProjectionImageVisible(image, regionElement, styles))
             .Select(image => new {
                 Image = image,
                 RenderKey = DescribeImageSource(GetImageSourceKey(image))
             })
            .ToList();
        var sourceByRenderKey = sourceItems.ToDictionary(
            item => item.RenderKey, item => item.Image, StringComparer.Ordinal);
        var ordered = new List<IHtmlImageElement>();
        var retainedKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach ((HtmlRenderImage image, double _) in EnumerateImages(
                     region.Visuals, includeBackgroundImages: false)) {
            string renderKey = image.Source ?? string.Empty;
            if (retainedKeys.Add(renderKey)
                && sourceByRenderKey.TryGetValue(renderKey, out IHtmlImageElement? sourceImage)) {
                ordered.Add(sourceImage);
            }
        }

        // External resources intentionally render as placeholders during the synchronous geometry pass.
        // Keep their source elements after the renderer-ordered images so the destination's async resolver
        // can still fetch and embed them. Hidden descendants remain excluded.
        foreach (var item in sourceItems) {
            if (retainedKeys.Add(item.RenderKey)) ordered.Add(item.Image);
        }
        return ordered.AsReadOnly();
    }

    private static bool IsProjectionImageVisible(
        IHtmlImageElement image,
        IElement regionElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) =>
        IsProjectionElementVisible(image, regionElement, styles);

    private static bool IsProjectionElementVisible(
        IElement element,
        IElement regionElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (styles.TryGetValue(current, out HtmlComputedStyle? style)) {
                string display = style.GetValue("display").Trim();
                string visibility = style.GetValue("visibility").Trim();
                if (display.Equals("none", StringComparison.OrdinalIgnoreCase)
                    || visibility.Equals("hidden", StringComparison.OrdinalIgnoreCase)
                    || visibility.Equals("collapse", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
            }
            if (ReferenceEquals(current, regionElement)) break;
        }
        return true;
    }

    private static int TryGetOwningElementNumber(IElement element, IReadOnlyList<IElement> owners) {
        for (int index = 0; index < owners.Count; index++) {
            if (ReferenceEquals(owners[index], element) || owners[index].Contains(element)) return index + 1;
        }
        return 0;
    }

    private static bool DiagnosticsAreEquivalent(HtmlDiagnostic first, HtmlDiagnostic second) =>
        string.Equals(first.Component, second.Component, StringComparison.Ordinal)
        && string.Equals(first.Code, second.Code, StringComparison.Ordinal)
        && string.Equals(first.Message, second.Message, StringComparison.Ordinal)
        && first.Severity == second.Severity
        && string.Equals(first.Source, second.Source, StringComparison.Ordinal)
        && string.Equals(first.Detail, second.Detail, StringComparison.Ordinal)
        && first.LossKind == second.LossKind;

    private static IEnumerable<EditableLayoutRegionOccurrence> EnumerateRegions(
        IEnumerable<HtmlRenderVisual> visuals,
        int pageNumber,
        (double X, double Y)? sectionOrigin,
        (double X, double Y)? tableOrigin = null) {
        foreach (HtmlRenderVisual visual in visuals) {
            (double X, double Y)? childSectionOrigin = sectionOrigin == null
                && visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Section } section
                    ? (section.X, section.Y)
                    : sectionOrigin;
            (double X, double Y)? childTableOrigin = tableOrigin == null
                && visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Table } table
                    ? (table.X, table.Y)
                    : tableOrigin;
            if (visual is HtmlRenderLayoutRegion region) {
                yield return new EditableLayoutRegionOccurrence(
                    pageNumber, region,
                    childSectionOrigin?.X ?? 0D, childSectionOrigin?.Y ?? 0D,
                    childTableOrigin?.X ?? 0D, childTableOrigin?.Y ?? 0D);
            }
            foreach (EditableLayoutRegionOccurrence child in EnumerateRegions(
                         GetChildren(visual), pageNumber, childSectionOrigin, childTableOrigin)) yield return child;
        }
    }

    private readonly struct EditableLayoutRegionOccurrence {
        internal EditableLayoutRegionOccurrence(
            int page,
            HtmlRenderLayoutRegion region,
            double sectionOriginX,
            double sectionOriginY,
            double tableOriginX,
            double tableOriginY) {
            Page = page;
            Region = region;
            SectionOriginX = sectionOriginX;
            SectionOriginY = sectionOriginY;
            TableOriginX = tableOriginX;
            TableOriginY = tableOriginY;
        }

        internal int Page { get; }
        internal HtmlRenderLayoutRegion Region { get; }
        internal double SectionOriginX { get; }
        internal double SectionOriginY { get; }
        internal double TableOriginX { get; }
        internal double TableOriginY { get; }
    }

    private static IEnumerable<HtmlRenderVisual> GetChildren(HtmlRenderVisual visual) => visual switch {
        HtmlRenderLayoutRegion region => region.Visuals,
        HtmlRenderSemanticGroup group => group.Visuals,
        HtmlRenderLogicalTextGroup group => group.Visuals,
        HtmlRenderEffectGroup group => group.Visuals,
        HtmlRenderClipGroup group => group.Visuals,
        HtmlRenderPathClipGroup group => group.Visuals,
        HtmlRenderFormField field => field.Visuals,
        _ => Array.Empty<HtmlRenderVisual>()
    };
}