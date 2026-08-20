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
        IReadOnlyList<HtmlDiagnostic> diagnostics) {
        RemainingDocument = remainingDocument;
        RenderedDocument = renderedDocument;
        Regions = regions;
        _sourceImages = sourceImages;
        Diagnostics = diagnostics;
    }

    private readonly IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> _sourceImages;

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
}

/// <summary>Creates one shared editable-layout plan for DOCX, RTF, XLSX, and PPTX adapters.</summary>
public static class HtmlEditableLayoutProjector {
    internal const string RegionAttribute = "data-officeimo-editable-layout-region";
    internal const string ImageAttribute = "data-officeimo-editable-layout-image";
    private const string ImageSourcePrefix = "img[officeimo-layout-image=";
    private static readonly HashSet<string> SemanticRichElementNames = new(StringComparer.OrdinalIgnoreCase) {
        "a", "abbr", "b", "blockquote", "button", "cite", "code", "dd", "del", "details", "dfn", "dl", "dt",
        "em", "fieldset", "form", "h1", "h2", "h3", "h4", "h5", "h6", "i", "input", "ins", "kbd", "label",
        "li", "mark", "ol", "pre", "q", "s", "samp", "select", "strong", "sub", "summary", "sup", "table",
        "textarea", "time", "u", "ul", "var"
    };

    /// <summary>Projects bounded positioned, floating, flex, and grid regions through the managed layout engine.</summary>
    public static HtmlEditableLayoutProjection Project(
        HtmlConversionDocument document,
        HtmlRenderOptions? renderOptions = null,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen,
        HtmlEditableLayoutRegionKinds regionKinds = HtmlEditableLayoutRegionKinds.All,
        int? maximumEditableSurfaceNumber = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if ((regionKinds & ~HtmlEditableLayoutRegionKinds.All) != 0) throw new ArgumentOutOfRangeException(nameof(regionKinds));
        if (maximumEditableSurfaceNumber < 0) throw new ArgumentOutOfRangeException(nameof(maximumEditableSurfaceNumber));
        IHtmlDocument adapterDocument = document.CreateSourceDocumentForConversion();
        foreach (IElement element in adapterDocument.QuerySelectorAll("[" + RegionAttribute + "]").ToArray()) {
            element.RemoveAttribute(RegionAttribute);
        }
        IReadOnlyList<HtmlGenericSectionProjection> semanticSections =
            HtmlGenericDocumentProjector.CreateSections(adapterDocument);
        IReadOnlyList<IElement> semanticTables = HtmlGenericDocumentProjector.SelectRootTables(adapterDocument);
        HtmlRenderOptions options = renderOptions?.Clone() ?? new HtmlRenderOptions();
        options.Mode = mediaContext == HtmlCssMediaContext.Print
            ? HtmlRenderMode.Paged
            : HtmlRenderMode.Continuous;
        IReadOnlyList<IHtmlStyleElement> callerStyles = HtmlRenderAdditionalStylesheetApplier.Apply(
            adapterDocument, options.AdditionalStylesheets.ToList());
        options.AdditionalStylesheets.Clear();
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(
            adapterDocument,
            mediaContext,
            document.Limits);
        int key = 0;
        var candidateElements = new Dictionary<string, IElement>(StringComparer.Ordinal);
        var semanticRichCandidates = new List<IElement>();
        foreach (IElement element in adapterDocument.QuerySelectorAll("*")) {
            if (!styles.TryGetValue(element, out HtmlComputedStyle? style)
                || !TryGetCandidateKind(element, style, out HtmlEditableLayoutRegionKinds candidateKind)
                || (regionKinds & candidateKind) == 0) continue;
            if (ContainsSemanticRichContent(element)) {
                semanticRichCandidates.Add(element);
                continue;
            }
            string sourceKey = (++key).ToString(System.Globalization.CultureInfo.InvariantCulture);
            element.SetAttribute(RegionAttribute, sourceKey);
            candidateElements[sourceKey] = element;
        }
        int imageKey = 0;
        foreach (IElement image in adapterDocument.QuerySelectorAll("img")) {
            image.SetAttribute(ImageAttribute, (++imageKey).ToString(
                System.Globalization.CultureInfo.InvariantCulture));
        }

        options.EnableEditableLayoutRegions = true;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(adapterDocument, options);
        var occurrences = new List<(int Page, HtmlRenderLayoutRegion Region)>();
        foreach (HtmlRenderPage page in rendered.Pages) {
            foreach (HtmlRenderLayoutRegion region in EnumerateRegions(page.Scene)) occurrences.Add((page.PageNumber, region));
        }

        var diagnostics = new HtmlDiagnosticReport();
        foreach (IElement element in semanticRichCandidates) {
            diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                "An editable layout region stayed in semantic flow so rich document content would not be flattened.",
                HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
                "semanticContent=true", OfficeConversionLossKind.Approximation);
        }
        var accepted = new List<HtmlRenderLayoutRegion>();
        foreach (IGrouping<string, (int Page, HtmlRenderLayoutRegion Region)> group in occurrences.GroupBy(item => item.Region.SourceKey)) {
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
            if (maximumEditableSurfaceNumber.HasValue
                && selected.SurfaceNumber > maximumEditableSurfaceNumber.Value) {
                diagnostics.Add("OfficeIMO.Html", HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "An editable layout region stayed in semantic flow because its rendered surface cannot be mapped natively by this destination.",
                    HtmlDiagnosticSeverity.Warning, selected.Source,
                    "surface=" + selected.SurfaceNumber + "; maximumEditableSurface=" + maximumEditableSurfaceNumber.Value,
                    OfficeConversionLossKind.Approximation);
                continue;
            }
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
        IReadOnlyDictionary<string, IReadOnlyList<IHtmlImageElement>> sourceImages = accepted.ToDictionary(
            region => region.SourceKey,
            region => {
                var renderedSources = new HashSet<string>(
                    EnumerateImages(region.Visuals, includeBackgroundImages: false)
                        .Select(item => item.Image.Source ?? string.Empty),
                    StringComparer.Ordinal);
                return (IReadOnlyList<IHtmlImageElement>)candidateElements[region.SourceKey]
                    .QuerySelectorAll("img")
                    .OfType<IHtmlImageElement>()
                    .Where(image => renderedSources.Contains(DescribeImageSource(
                        image.GetAttribute(ImageAttribute))))
                    .ToList()
                    .AsReadOnly();
            },
            StringComparer.Ordinal);
        foreach (IHtmlImageElement image in sourceImages.Values.SelectMany(images => images)) {
            image.RemoveAttribute(ImageAttribute);
        }
        foreach (IElement element in adapterDocument.QuerySelectorAll("[" + RegionAttribute + "]").ToArray()) {
            string? sourceKey = element.GetAttribute(RegionAttribute);
            if (sourceKey != null && acceptedKeys.Contains(sourceKey)) element.Remove();
            else element.RemoveAttribute(RegionAttribute);
        }
        foreach (IElement image in adapterDocument.QuerySelectorAll("[" + ImageAttribute + "]").ToArray()) {
            image.RemoveAttribute(ImageAttribute);
        }
        foreach (IHtmlStyleElement style in callerStyles) style.Remove();
        return new HtmlEditableLayoutProjection(adapterDocument, rendered, accepted.AsReadOnly(), sourceImages, diagnostics.Diagnostics);
    }

    internal static HtmlSemanticDocument BuildRemainingSemanticDocument(
        HtmlEditableLayoutProjection projection,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits limits) =>
        HtmlSemanticDocumentBuilder.FromDocument(projection.RemainingDocument, mediaContext, limits);

    internal static bool MayContainEditableLayoutRegions(
        HtmlConversionDocument document,
        HtmlEditableLayoutRegionKinds regionKinds) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        string html = document.SourceHtml;
        return ((regionKinds & HtmlEditableLayoutRegionKinds.Positioned) != 0
                && html.IndexOf("position", StringComparison.OrdinalIgnoreCase) >= 0)
            || ((regionKinds & HtmlEditableLayoutRegionKinds.Floating) != 0
                && html.IndexOf("float", StringComparison.OrdinalIgnoreCase) >= 0)
            || ((regionKinds & (HtmlEditableLayoutRegionKinds.Flex | HtmlEditableLayoutRegionKinds.Grid)) != 0
                && html.IndexOf("display", StringComparison.OrdinalIgnoreCase) >= 0);
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
            string? sourceKey = ancestor.GetAttribute(RegionAttribute);
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

    private static bool IsBackgroundImage(HtmlRenderImage image) =>
        image.Source?.IndexOf(":background-image", StringComparison.Ordinal) >= 0;

    internal static string DescribeImageSource(string? key) =>
        ImageSourcePrefix + (key ?? string.Empty) + "]";

    private static int TryGetOwningElementNumber(IElement element, IReadOnlyList<IElement> owners) {
        for (int index = 0; index < owners.Count; index++) {
            if (ReferenceEquals(owners[index], element) || owners[index].Contains(element)) return index + 1;
        }
        return 0;
    }

    private static bool ContainsSemanticRichContent(IElement element) {
        if (SemanticRichElementNames.Contains(element.LocalName)) return true;
        return element.QuerySelectorAll("*").Any(child => SemanticRichElementNames.Contains(child.LocalName));
    }

    private static IEnumerable<HtmlRenderLayoutRegion> EnumerateRegions(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderLayoutRegion region) yield return region;
            foreach (HtmlRenderLayoutRegion child in EnumerateRegions(GetChildren(visual))) yield return child;
        }
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
