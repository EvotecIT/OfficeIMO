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
        int maximumCharacters,
        int maximumNodes,
        int maximumDepth,
        Uri? baseUri,
        Func<Uri, string?>? stylesheetResolver,
        CancellationToken token) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(target);
        token.ThrowIfCancellationRequested();
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxInputCharacters = maximumCharacters;
        limits.MaxHtmlNodes = maximumNodes;
        limits.MaxHtmlDepth = maximumDepth;
        HtmlConversionInputGuard.ValidateDocument(document, limits, token);
        IElement[] sourceElements = document.QuerySelectorAll("*").ToArray();
        int targetIndex = Array.IndexOf(sourceElements, target);
        if (targetIndex < 0) return HtmlInteractionLayoutResult.Detached;
        Uri?[] stylesheetUrls = document.QuerySelectorAll("link").OfType<IHtmlLinkElement>()
            .Select(link => ResolveStylesheetUrl(link, baseUri))
            .ToArray();

        string html = document.DocumentElement?.OuterHtml ?? string.Empty;
        if (html.Length > maximumCharacters) throw new ArgumentException("The live document exceeds the interaction layout input budget.");
        IHtmlDocument clone = HtmlDocumentParser.ParseDocument("<!doctype html><html><head></head><body></body></html>", token);
        if (document.DocumentElement != null && clone.DocumentElement != null) {
            var importedRoot = (IElement)clone.Import(document.DocumentElement, deep: true);
            NativeFormState.CopyTree(document.DocumentElement, importedRoot, token);
            HydrateStylesheets(clone, importedRoot, stylesheetUrls, stylesheetResolver, token);
            clone.ReplaceChild(importedRoot, clone.DocumentElement);
        }
        IElement[] clonedElements = clone.QuerySelectorAll("*").ToArray();
        if (targetIndex >= clonedElements.Length) return HtmlInteractionLayoutResult.Detached;
        IElement clonedTarget = clonedElements[targetIndex];
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = viewportWidth,
            ViewportHeight = viewportHeight,
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
        if (!cssVisible) return new HtmlInteractionLayoutResult(true, false, acceptsPointerEvents, null, viewportWidth, viewportHeight);

        string marker = "officeimo-interaction:" + Guid.NewGuid().ToString("N");
        HtmlRenderSourceIdentity.Register(clonedTarget, marker);
        if (string.Equals(targetStyle!.GetValue("display").Trim(), "inline", StringComparison.OrdinalIgnoreCase)) {
            // A pure inline box is represented by its descendant runs rather than a
            // standalone block visual. Mark that subtree so wrapped text fragments
            // contribute to this element's union without widening block/atomic boxes.
            foreach (IElement descendant in clonedTarget.QuerySelectorAll("*")) {
                HtmlRenderSourceIdentity.Register(descendant, marker);
            }
        }
        var diagnostics = new HtmlDiagnosticReport();
        HtmlRenderDocument rendered = new HtmlRenderLayoutEngine(clone, styles, options, diagnostics, cancellationToken: token).Render();
        HtmlInteractionRect? bounds = Bounds(rendered.Pages.SelectMany(page => Enumerate(page.Scene)), marker);
        double documentWidth = rendered.Pages.Max(page => page.Width);
        double documentHeight = rendered.Pages.Sum(page => page.Height);
        return new HtmlInteractionLayoutResult(true, bounds.HasValue, acceptsPointerEvents, bounds, documentWidth, documentHeight);
    }

    private static void HydrateStylesheets(IHtmlDocument document, IParentNode root, IReadOnlyList<Uri?> urls,
        Func<Uri, string?>? resolver, CancellationToken token) {
        if (resolver == null) return;
        IHtmlLinkElement[] links = root.QuerySelectorAll("link").OfType<IHtmlLinkElement>().ToArray();
        for (int index = 0; index < links.Length && index < urls.Count; index++) {
            token.ThrowIfCancellationRequested();
            IHtmlLinkElement link = links[index];
            Uri? url = urls[index];
            if (url == null || link.IsDisabled || IsAlternateStylesheet(link)) continue;
            string? css = resolver(url);
            if (css == null || link.Parent == null) continue;
            var style = (IHtmlStyleElement)document.CreateElement("style");
            style.TextContent = css;
            if (link.HasAttribute("media")) style.SetAttribute("media", link.GetAttribute("media"));
            link.Parent.ReplaceChild(style, link);
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

    private static HtmlInteractionRect? Bounds(IEnumerable<HtmlRenderVisual> visuals, string marker) {
        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual.Source == null || !visual.Source.StartsWith(marker, StringComparison.Ordinal)) continue;
            left = Math.Min(left, visual.X);
            top = Math.Min(top, visual.Y);
            right = Math.Max(right, visual.X + visual.Width);
            bottom = Math.Max(bottom, visual.Y + visual.Height);
        }
        return double.IsPositiveInfinity(left) || right <= left || bottom <= top
            ? null
            : new HtmlInteractionRect(left, top, right - left, bottom - top);
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
}

internal readonly record struct HtmlInteractionRect(double X, double Y, double Width, double Height);

internal readonly record struct HtmlInteractionLayoutResult(
    bool IsConnected,
    bool HasLayoutBox,
    bool AcceptsPointerEvents,
    HtmlInteractionRect? Bounds,
    double DocumentWidth,
    double DocumentHeight) {
    internal static HtmlInteractionLayoutResult Detached => new(false, false, false, null, 0D, 0D);
}
