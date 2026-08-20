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
public static class HtmlEditableLayoutProjector {
    internal const string RegionAttribute = "data-officeimo-editable-layout-region";
    internal const string ImageAttribute = "data-officeimo-editable-layout-image";
    private const string ImageSourcePrefix = "img[officeimo-layout-image=";
    private static readonly HashSet<string> SemanticRichElementNames = new(StringComparer.OrdinalIgnoreCase) {
        "a", "abbr", "audio", "b", "blockquote", "br", "button", "canvas", "cite", "code", "dd", "del",
        "details", "dfn", "dl", "dt", "em", "embed", "fieldset", "figure", "figcaption", "form", "h1", "h2",
        "h3", "h4", "h5", "h6", "hr", "i", "iframe", "input", "ins", "kbd", "label", "li", "mark",
        "object", "ol", "p", "picture", "pre", "q", "s", "samp", "select", "strong", "sub", "summary",
        "sup", "svg", "table", "textarea", "time", "u", "ul", "var", "video"
    };
    private static readonly string[] RichTextStyleProperties = {
        "color", "direction", "font-family", "font-size", "font-style", "font-variant", "font-weight",
        "letter-spacing", "line-height",
        "text-decoration", "text-decoration-color", "text-decoration-line", "text-decoration-style", "text-shadow",
        "text-transform", "unicode-bidi", "vertical-align", "white-space", "word-spacing"
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
            if (ContainsSemanticRichContent(element, styles)) {
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
            region => CreateOrderedSourceImages(region, candidateElements[region.SourceKey], styles),
            StringComparer.Ordinal);
        IReadOnlyDictionary<string, IHtmlImageElement> sourceImagesByRenderKey = sourceImages.Values
            .SelectMany(images => images)
            .ToDictionary(image => DescribeImageSource(image.GetAttribute(ImageAttribute)), image => image,
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
        IReadOnlyList<HtmlDiagnostic> projectionDiagnostics = rendered.Diagnostics
            .Where(renderDiagnostic => !document.Diagnostics.Any(sourceDiagnostic =>
                DiagnosticsAreEquivalent(sourceDiagnostic, renderDiagnostic)))
            .Concat(diagnostics.Diagnostics)
            .ToArray();
        return new HtmlEditableLayoutProjection(adapterDocument, rendered, accepted.AsReadOnly(), sourceImages,
            sourceImagesByRenderKey, projectionDiagnostics);
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

    private static IReadOnlyList<IHtmlImageElement> CreateOrderedSourceImages(
        HtmlRenderLayoutRegion region,
        IElement regionElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        var sourceItems = regionElement.QuerySelectorAll("img")
            .OfType<IHtmlImageElement>()
            .Where(image => IsProjectionImageVisible(image, regionElement, styles))
            .Select(image => new {
                Image = image,
                RenderKey = DescribeImageSource(image.GetAttribute(ImageAttribute))
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
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        for (IElement? current = image; current != null; current = current.ParentElement) {
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

    private static bool ContainsSemanticRichContent(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        if (string.Equals(element.LocalName, "img", StringComparison.OrdinalIgnoreCase)) return false;
        if (SemanticRichElementNames.Contains(element.LocalName)) return true;
        if (HasDistinctRichTextStyle(element, styles)) return true;
        return element.QuerySelectorAll("*").Any(child => {
            if (string.Equals(child.LocalName, "img", StringComparison.OrdinalIgnoreCase)) return false;
            return SemanticRichElementNames.Contains(child.LocalName)
                || HasDistinctRichTextStyle(child, styles)
                || HasDistinctStyle(child, styles, RichDescendantVisualStyleProperties);
        });
    }

    private static bool HasDistinctRichTextStyle(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        return HasDistinctStyle(element, styles, RichTextStyleProperties);
    }

    private static bool HasDistinctStyle(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        IReadOnlyList<string> properties) {
        if (!styles.TryGetValue(element, out HtmlComputedStyle? style)
            || element.ParentElement == null
            || !styles.TryGetValue(element.ParentElement, out HtmlComputedStyle? parentStyle)) return false;
        return properties.Any(property => !style.IsInheritedValue(property)
            && !style.IsResetValue(property)
            && !string.IsNullOrWhiteSpace(style.GetValue(property))
            && !string.Equals(style.GetValue(property), parentStyle.GetValue(property),
                StringComparison.OrdinalIgnoreCase));
    }

    private static bool DiagnosticsAreEquivalent(HtmlDiagnostic first, HtmlDiagnostic second) =>
        string.Equals(first.Component, second.Component, StringComparison.Ordinal)
        && string.Equals(first.Code, second.Code, StringComparison.Ordinal)
        && string.Equals(first.Message, second.Message, StringComparison.Ordinal)
        && first.Severity == second.Severity
        && string.Equals(first.Source, second.Source, StringComparison.Ordinal)
        && string.Equals(first.Detail, second.Detail, StringComparison.Ordinal)
        && first.LossKind == second.LossKind;

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
