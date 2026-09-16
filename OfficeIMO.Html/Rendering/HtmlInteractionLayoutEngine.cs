using AngleSharp;
using AngleSharp.Css.Dom;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

// This is the shared static-layout bridge used by runtime actionability. It works
// on a clone, so measurement cannot mutate or notify the live scripted document.
internal static class HtmlInteractionLayoutEngine {
    internal static HtmlInteractionLayoutResult Measure(
        IHtmlDocument document,
        IElement target,
        double viewportWidth,
        double viewportHeight,
        double devicePixelRatio,
        double scrollX,
        double scrollY,
        int maximumCharacters,
        int maximumNodes,
        int maximumDepth,
        int maximumStylesheetImportDepth,
        Uri? baseUri,
        Func<Uri, string?>? stylesheetResolver,
        CancellationToken token) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (target == null) throw new ArgumentNullException(nameof(target));
        token.ThrowIfCancellationRequested();
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxInputCharacters = maximumCharacters;
        limits.MaxHtmlNodes = maximumNodes;
        limits.MaxHtmlDepth = maximumDepth;
        HtmlConversionInputGuard.ValidateDocument(document, limits, token);
        IElement[] sourceElements = document.QuerySelectorAll("*").ToArray();
        int targetIndex = Array.IndexOf(sourceElements, target);
        if (targetIndex < 0) return HtmlInteractionLayoutResult.Detached;
        var mediaFeatures = new HtmlRenderMediaFeatures {
            ResolutionDpi = devicePixelRatio * HtmlRenderOptions.CssPixelsPerInch
        };
        StylesheetSnapshot[] stylesheets = CaptureStylesheets(
            sourceElements, baseUri, stylesheetResolver, limits, viewportWidth, viewportHeight,
            mediaFeatures, maximumStylesheetImportDepth, token);

        string html = document.DocumentElement?.OuterHtml ?? string.Empty;
        if (html.Length > maximumCharacters) throw new ArgumentException("The live document exceeds the interaction layout input budget.");
        IHtmlDocument clone = HtmlDocumentParser.ParseDocument("<!doctype html><html><head></head><body></body></html>", token);
        IElement? clonedTarget = null;
        if (document.DocumentElement != null && clone.DocumentElement != null) {
            var importedRoot = (IElement)clone.Import(document.DocumentElement, deep: true);
            NativeFormState.CopyTree(document.DocumentElement, importedRoot, token);
            IElement[] importedElements = new[] { importedRoot }.Concat(importedRoot.QuerySelectorAll("*")).ToArray();
            if (targetIndex < importedElements.Length) clonedTarget = importedElements[targetIndex];
            HydrateStylesheets(clone, importedRoot, stylesheets, limits, viewportWidth, viewportHeight, mediaFeatures,
                maximumStylesheetImportDepth, stylesheetResolver, token);
            clone.ReplaceChild(importedRoot, clone.DocumentElement);
        }
        IElement[] clonedElements = clone.QuerySelectorAll("*").ToArray();
        if (clonedTarget == null) return HtmlInteractionLayoutResult.Detached;
        targetIndex = Array.IndexOf(clonedElements, clonedTarget);
        if (targetIndex < 0) return HtmlInteractionLayoutResult.Detached;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = viewportWidth,
            ViewportHeight = viewportHeight,
            MediaFeatures = mediaFeatures,
            Margins = HtmlRenderMargins.All(0D),
            BaseUri = baseUri ?? (Uri.TryCreate(document.Url, UriKind.Absolute, out Uri? documentUrl) ? documentUrl : null),
            MaxHtmlNodes = maximumNodes,
            MaxLayoutDepth = maximumDepth
        };
        options.Validate();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(clone, options, limits);
        bool cssVisible = CssVisible(clonedTarget, styles.Elements);
        bool acceptsPointerEvents = styles.Elements.TryGetValue(clonedTarget, out HtmlComputedStyle? targetStyle)
            && !string.Equals(targetStyle.GetValue("pointer-events").Trim(), "none", StringComparison.OrdinalIgnoreCase);
        if (!cssVisible) return new HtmlInteractionLayoutResult(true, false, acceptsPointerEvents, false, null, viewportWidth, viewportHeight);

        string marker = "officeimo-interaction:" + Guid.NewGuid().ToString("N") + ":";
        for (int index = 0; index < clonedElements.Length; index++)
            HtmlRenderSourceIdentity.Register(clonedElements[index], marker + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
        var hitTargetSources = new HashSet<int>();
        for (int index = 0; index < clonedElements.Length; index++)
            if (IsWithin(clonedElements[index], clonedTarget)) hitTargetSources.Add(index);
        HashSet<int> boundsSources = new() { targetIndex };
        if (string.Equals(targetStyle!.GetValue("display").Trim(), "inline", StringComparison.OrdinalIgnoreCase)) {
            // A pure inline box is represented by its descendant runs rather than a
            // standalone block visual. Mark that subtree so wrapped text fragments
            // contribute to this element's union without widening block/atomic boxes.
            boundsSources = hitTargetSources;
        }
        var diagnostics = new HtmlDiagnosticReport();
        HtmlRenderDocument rendered = new HtmlRenderLayoutEngine(clone, styles, options, diagnostics, cancellationToken: token).Render();
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => Enumerate(page.Scene)).ToArray();
        HtmlInteractionRect? bounds = Bounds(visuals, marker, boundsSources, clonedElements, styles.Elements, scrollX, scrollY);
        bool hasUnqualifiedScrolledSticky = (Math.Abs(scrollX) > 0.0001D || Math.Abs(scrollY) > 0.0001D)
            && clonedElements.Any(element => CssVisible(element, styles.Elements) && HasPosition(element, styles.Elements, "sticky"));
        bool receivesPointerAtCenter = !hasUnqualifiedScrolledSticky && bounds.HasValue
            && HitTest(visuals, marker, clonedElements, styles.Elements, scrollX, scrollY,
                bounds.Value.X + bounds.Value.Width / 2D, bounds.Value.Y + bounds.Value.Height / 2D) is int hit
            && hitTargetSources.Contains(hit);
        double documentWidth = rendered.Pages.Max(page => page.Width);
        double documentHeight = rendered.Pages.Sum(page => page.Height);
        return new HtmlInteractionLayoutResult(true, bounds.HasValue, acceptsPointerEvents, receivesPointerAtCenter, bounds, documentWidth, documentHeight);
    }

    private static StylesheetSnapshot[] CaptureStylesheets(
        IReadOnlyList<IElement> elements,
        Uri? baseUri,
        Func<Uri, string?>? resolver,
        HtmlConversionLimits limits,
        double viewportWidth,
        double viewportHeight,
        HtmlRenderMediaFeatures mediaFeatures,
        int maximumImportDepth,
        CancellationToken token) {
        var snapshots = new List<StylesheetSnapshot>();
        for (int index = 0; index < elements.Count; index++) {
            token.ThrowIfCancellationRequested();
            IElement element = elements[index];
            bool isInline = element is IHtmlStyleElement;
            bool isLink = element is IHtmlLinkElement link && IsStylesheet(link);
            if (!isInline && !isLink) continue;
            string type = element.GetAttribute("type") ?? string.Empty;
            if (type.Length != 0 && !type.Equals("text/css", StringComparison.OrdinalIgnoreCase)) continue;
            Uri? stylesheetUri = isLink ? ResolveStylesheetUrl((IHtmlLinkElement)element, baseUri) : baseUri;
            IStyleSheet? sheet = (element as ILinkStyle)?.Sheet;
            string? css = sheet is ICssStyleSheet cssSheet
                ? stylesheetUri == null ? cssSheet.ToCss() : SerializeLiveStylesheet(
                    cssSheet, stylesheetUri, limits, viewportWidth, viewportHeight, mediaFeatures,
                    maximumImportDepth, 0, new HashSet<ICssStyleSheet>(), token)
                : isInline ? element.TextContent ?? string.Empty
                : stylesheetUri != null ? resolver?.Invoke(stylesheetUri) : null;
            bool disabled = sheet?.IsDisabled == true
                || element is IHtmlStyleElement style && style.IsDisabled
                || element is IHtmlLinkElement stylesheetLink && stylesheetLink.IsDisabled;
            snapshots.Add(new StylesheetSnapshot(index, element.LocalName, css, stylesheetUri, disabled,
                element is IHtmlLinkElement alternate && IsAlternateStylesheet(alternate),
                element.GetAttribute("media")));
        }
        return snapshots.ToArray();
    }

    private static void HydrateStylesheets(
        IHtmlDocument document,
        IParentNode root,
        IReadOnlyList<StylesheetSnapshot> snapshots,
        HtmlConversionLimits limits,
        double viewportWidth,
        double viewportHeight,
        HtmlRenderMediaFeatures mediaFeatures,
        int maximumImportDepth,
        Func<Uri, string?>? resolver,
        CancellationToken token) {
        IElement[] elements = root is IElement rootElement
            ? new[] { rootElement }.Concat(root.QuerySelectorAll("*")).ToArray()
            : root.QuerySelectorAll("*").ToArray();
        var budget = new HtmlCssByteBudget(limits);
        foreach (StylesheetSnapshot snapshot in snapshots) {
            token.ThrowIfCancellationRequested();
            if (snapshot.ElementIndex < 0 || snapshot.ElementIndex >= elements.Length) continue;
            IElement source = elements[snapshot.ElementIndex];
            if (!source.LocalName.Equals(snapshot.LocalName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The cloned stylesheet element mapping changed during interaction layout preparation.");
            if (snapshot.Disabled || snapshot.IsAlternate || snapshot.Css == null) {
                source.Remove();
                continue;
            }
            budget.ReserveOrThrow(snapshot.Css);
            string css = snapshot.StylesheetUri == null || resolver == null
                ? snapshot.Css
                : ExpandStylesheetImports(snapshot.Css, snapshot.StylesheetUri, resolver, limits,
                    viewportWidth, viewportHeight, mediaFeatures, maximumImportDepth, 0,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase), budget, token);
            IHtmlStyleElement style = source as IHtmlStyleElement ?? (IHtmlStyleElement)document.CreateElement("style");
            style.TextContent = css;
            if (!string.IsNullOrWhiteSpace(snapshot.Media)) style.SetAttribute("media", snapshot.Media);
            if (!ReferenceEquals(style, source) && source.Parent != null) source.Parent.ReplaceChild(style, source);
        }
    }

    private static string ExpandStylesheetImports(
        string css,
        Uri stylesheetUri,
        Func<Uri, string?> resolver,
        HtmlConversionLimits limits,
        double viewportWidth,
        double viewportHeight,
        HtmlRenderMediaFeatures mediaFeatures,
        int maximumDepth,
        int depth,
        HashSet<string> active,
        HtmlCssByteBudget budget,
        CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string key = stylesheetUri.AbsoluteUri;
        if (!active.Add(key)) return string.Empty;
        try {
            var resourceOptions = new HtmlResourcePipelineOptions {
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                Limits = limits.Clone(),
                MaxResponsiveImageCandidates = limits.MaxResponsiveImageCandidates,
                MaxResponsiveImageSizesCharacters = limits.MaxResponsiveImageSizesCharacters,
                MediaContext = HtmlCssMediaContext.Screen,
                MediaWidth = viewportWidth,
                MediaHeight = viewportHeight,
                DevicePixelRatio = mediaFeatures.ResolutionDpi / HtmlRenderOptions.CssPixelsPerInch,
                MediaFeatures = mediaFeatures
            };
            HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(css, stylesheetUri, resourceOptions);
            var builder = new System.Text.StringBuilder(analysis.Css);
            for (int index = analysis.Imports.Count - 1; index >= 0; index--) {
                token.ThrowIfCancellationRequested();
                HtmlExternalStylesheetImport import = analysis.Imports[index];
                string replacement = string.Empty;
                if (depth < maximumDepth && import.IsApplicable && import.Reference.IsAllowed
                    && Uri.TryCreate(import.Reference.ResolvedSource, UriKind.Absolute, out Uri? importedUri)
                    && !active.Contains(importedUri.AbsoluteUri)) {
                    string? importedCss = resolver(importedUri);
                    if (importedCss != null) {
                        budget.ReserveOrThrow(importedCss);
                        replacement = ExpandStylesheetImports(importedCss, importedUri, resolver, limits,
                            viewportWidth, viewportHeight, mediaFeatures, maximumDepth, depth + 1, active, budget, token);
                    }
                }
                builder.Remove(import.Start, import.End - import.Start);
                builder.Insert(import.Start, replacement);
            }
            return HtmlResourcePipeline.RebaseExternalStylesheetUrls(
                builder.ToString(), stylesheetUri, HtmlUrlPolicy.CreateWebResourceProfile());
        } finally {
            active.Remove(key);
        }
    }

    private static string SerializeLiveStylesheet(
        ICssStyleSheet sheet,
        Uri stylesheetUri,
        HtmlConversionLimits limits,
        double viewportWidth,
        double viewportHeight,
        HtmlRenderMediaFeatures mediaFeatures,
        int maximumDepth,
        int depth,
        HashSet<ICssStyleSheet> active,
        CancellationToken token) {
        if (!active.Add(sheet)) return string.Empty;
        try {
            var builder = new System.Text.StringBuilder();
            var resourceOptions = new HtmlResourcePipelineOptions {
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                Limits = limits.Clone(),
                MaxResponsiveImageCandidates = limits.MaxResponsiveImageCandidates,
                MaxResponsiveImageSizesCharacters = limits.MaxResponsiveImageSizesCharacters,
                MediaContext = HtmlCssMediaContext.Screen,
                MediaWidth = viewportWidth,
                MediaHeight = viewportHeight,
                DevicePixelRatio = mediaFeatures.ResolutionDpi / HtmlRenderOptions.CssPixelsPerInch,
                MediaFeatures = mediaFeatures
            };
            foreach (ICssRule rule in sheet.Rules) {
                token.ThrowIfCancellationRequested();
                string ruleCss = rule.ToCss();
                if (rule is not ICssImportRule import) {
                    builder.AppendLine(ruleCss);
                    continue;
                }
                HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(ruleCss, stylesheetUri, resourceOptions);
                if (analysis.Imports.Count == 0 || !analysis.Imports[0].IsApplicable || depth >= maximumDepth) continue;
                if (import.Sheet is not ICssStyleSheet importedSheet) {
                    builder.AppendLine(ruleCss);
                    continue;
                }
                Uri? importedUri = Uri.TryCreate(importedSheet.Href, UriKind.Absolute, out Uri? absolute)
                    ? absolute
                    : Uri.TryCreate(stylesheetUri, import.Href, out Uri? relative) ? relative : null;
                if (importedUri == null || !importedUri.IsAbsoluteUri || importedSheet.IsDisabled) continue;
                builder.AppendLine(SerializeLiveStylesheet(importedSheet, importedUri, limits,
                    viewportWidth, viewportHeight, mediaFeatures, maximumDepth, depth + 1, active, token));
            }
            return HtmlResourcePipeline.RebaseExternalStylesheetUrls(
                builder.ToString(), stylesheetUri, HtmlUrlPolicy.CreateWebResourceProfile());
        } finally {
            active.Remove(sheet);
        }
    }

    private static bool IsStylesheet(IHtmlLinkElement link) => (link.GetAttribute("rel") ?? string.Empty)
        .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
        .Contains("stylesheet", StringComparer.OrdinalIgnoreCase);

    private static bool IsAlternateStylesheet(IHtmlLinkElement link) => (link.GetAttribute("rel") ?? string.Empty)
        .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
        .Contains("alternate", StringComparer.OrdinalIgnoreCase);

    private static Uri? ResolveStylesheetUrl(IHtmlLinkElement link, Uri? baseUri) {
        string? href = link.GetAttribute("href");
        return IsStylesheet(link) && !string.IsNullOrWhiteSpace(href)
            && Uri.TryCreate(baseUri, href, out Uri? resolved) && resolved.IsAbsoluteUri ? resolved : null;
    }

    private static bool CssVisible(IElement target, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        if (!styles.TryGetValue(target, out HtmlComputedStyle? targetStyle)) return false;
        string visibility = targetStyle.GetValue("visibility").Trim();
        if (visibility.Equals("hidden", StringComparison.OrdinalIgnoreCase)
            || visibility.Equals("collapse", StringComparison.OrdinalIgnoreCase)) return false;
        for (IElement? current = target; current != null; current = current.ParentElement) {
            if (styles.TryGetValue(current, out HtmlComputedStyle? style)
                && style.GetValue("display").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) return false;
        }
        return true;
    }

    private static HtmlInteractionRect? Bounds(
        IEnumerable<HtmlRenderVisual> visuals,
        string marker,
        HashSet<int> sources,
        IReadOnlyList<IElement> elements,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        double scrollX,
        double scrollY) {
        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (HtmlRenderVisual visual in visuals) {
            if (!TrySourceIndex(visual.Source, marker, out int source) || !sources.Contains(source)
                || source < 0 || source >= elements.Count) continue;
            PositionOffset(elements[source], styles, scrollX, scrollY, out double offsetX, out double offsetY);
            left = Math.Min(left, visual.X + offsetX);
            top = Math.Min(top, visual.Y + offsetY);
            right = Math.Max(right, visual.X + offsetX + visual.Width);
            bottom = Math.Max(bottom, visual.Y + offsetY + visual.Height);
        }
        return double.IsPositiveInfinity(left) || right <= left || bottom <= top
            ? null
            : new HtmlInteractionRect(left, top, right - left, bottom - top);
    }

    private static int? HitTest(
        IEnumerable<HtmlRenderVisual> visuals,
        string marker,
        IReadOnlyList<IElement> elements,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        double scrollX,
        double scrollY,
        double x,
        double y) {
        int? hit = null;
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup or HtmlRenderLogicalTextGroup or HtmlRenderClipGroup
                or HtmlRenderPathClipGroup or HtmlRenderEffectGroup) continue;
            if (!TrySourceIndex(visual.Source, marker, out int source) || source < 0 || source >= elements.Count) continue;
            IElement element = elements[source];
            PositionOffset(element, styles, scrollX, scrollY, out double offsetX, out double offsetY);
            if (x < visual.X + offsetX || y < visual.Y + offsetY
                || x >= visual.X + offsetX + visual.Width || y >= visual.Y + offsetY + visual.Height) continue;
            if (!CssVisible(element, styles) || !styles.TryGetValue(element, out HtmlComputedStyle? style)
                || style.GetValue("pointer-events").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) continue;
            hit = source;
        }
        return hit;
    }

    private static void PositionOffset(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        double scrollX,
        double scrollY,
        out double offsetX,
        out double offsetY) {
        bool fixedPosition = HasPosition(element, styles, "fixed");
        offsetX = fixedPosition ? scrollX : 0D;
        offsetY = fixedPosition ? scrollY : 0D;
    }

    private static bool HasPosition(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        string position) {
        for (IElement? current = element; current != null; current = current.ParentElement)
            if (styles.TryGetValue(current, out HtmlComputedStyle? style)
                && style.GetValue("position").Trim().Equals(position, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static bool TrySourceIndex(string? source, string marker, out int index) {
        index = -1;
        if (source == null || !source.StartsWith(marker, StringComparison.Ordinal)) return false;
        string suffix = source.Substring(marker.Length);
        int end = suffix.IndexOf(':');
        if (end >= 0) suffix = suffix.Substring(0, end);
        return int.TryParse(suffix, System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture, out index);
    }

    private static bool IsWithin(IElement candidate, IElement ancestor) {
        for (IElement? current = candidate; current != null; current = current.ParentElement)
            if (ReferenceEquals(current, ancestor)) return true;
        return false;
    }

    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLayoutRegion region ? region.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals
                : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }

    private readonly record struct StylesheetSnapshot(
        int ElementIndex,
        string LocalName,
        string? Css,
        Uri? StylesheetUri,
        bool Disabled,
        bool IsAlternate,
        string? Media);

}

internal readonly record struct HtmlInteractionRect(double X, double Y, double Width, double Height);

internal readonly record struct HtmlInteractionLayoutResult(
    bool IsConnected,
    bool HasLayoutBox,
    bool AcceptsPointerEvents,
    bool ReceivesPointerAtCenter,
    HtmlInteractionRect? Bounds,
    double DocumentWidth,
    double DocumentHeight) {
    internal static HtmlInteractionLayoutResult Detached => new(false, false, false, false, null, 0D, 0D);
}
