namespace OfficeIMO.Markdown;

internal static class AssetFactory {
    internal static System.Collections.Generic.IEnumerable<MarkdownHtmlAsset> PrismAssets(
        MarkdownPrismOptions po,
        MarkdownAssetMode mode,
        MarkdownCssDelivery cssDelivery,
        string? scopeSelector,
        MarkdownExternalTextResolver? resolver) {
        // Build URLs
        string coreJs = po.CdnBase.TrimEnd('/') + "/components/prism-core.min.js";
        string prismLightCss = po.CdnBase.TrimEnd('/') + "/themes/prism.min.css";
        string prismDarkCss = po.CdnBase.TrimEnd('/') + "/themes/prism-okaidia.min.css";

        // Theme CSS handling
        if (po.Theme == MarkdownPrismTheme.GithubAuto) {
            if (cssDelivery == MarkdownCssDelivery.LinkHref && mode == MarkdownAssetMode.Online) {
                var light = new MarkdownHtmlAsset("prism-theme:light", MarkdownHtmlAssetKind.Css, prismLightCss, null) { Media = "(prefers-color-scheme: light)" };
                var dark = new MarkdownHtmlAsset("prism-theme:dark", MarkdownHtmlAssetKind.Css, prismDarkCss, null) { Media = "(prefers-color-scheme: dark)" };
                yield return light; yield return dark;
            } else {
                var lightCss = HtmlRenderer.ResolveExternalText(prismLightCss, resolver);
                var darkCss = HtmlRenderer.ResolveExternalText(prismDarkCss, resolver);
                if (!string.IsNullOrEmpty(lightCss)) lightCss = HtmlRenderer.ScopeCss(lightCss, scopeSelector);
                if (!string.IsNullOrEmpty(darkCss)) darkCss = "@media (prefers-color-scheme: dark){\n" + HtmlRenderer.ScopeCss(darkCss, scopeSelector) + "\n}";
                yield return new MarkdownHtmlAsset("prism-theme:auto", MarkdownHtmlAssetKind.Css, null, (lightCss ?? string.Empty) + (string.IsNullOrEmpty(lightCss) || string.IsNullOrEmpty(darkCss) ? string.Empty : "\n") + (darkCss ?? string.Empty));
            }
        } else {
            string themeCss = po.Theme switch {
                MarkdownPrismTheme.Okaidia => prismDarkCss,
                MarkdownPrismTheme.GithubDark => prismDarkCss,
                _ => prismLightCss
            };
            if (cssDelivery == MarkdownCssDelivery.LinkHref && mode == MarkdownAssetMode.Online) {
                yield return new MarkdownHtmlAsset($"prism-theme:{po.Theme}", MarkdownHtmlAssetKind.Css, themeCss, null);
            } else {
                var css = HtmlRenderer.ResolveExternalText(themeCss, resolver);
                if (!string.IsNullOrEmpty(css)) css = HtmlRenderer.ScopeCss(css, scopeSelector);
                yield return new MarkdownHtmlAsset($"prism-theme:{po.Theme}", MarkdownHtmlAssetKind.Css, null, css);
            }
        }

        // Core JS
        if (mode == MarkdownAssetMode.Online) yield return new MarkdownHtmlAsset("prism-core", MarkdownHtmlAssetKind.Js, coreJs, null);
        else yield return new MarkdownHtmlAsset("prism-core", MarkdownHtmlAssetKind.Js, null, HtmlRenderer.ResolveExternalText(coreJs, resolver));

        // Languages
        foreach (var lang in po.Languages) {
            string url = po.CdnBase.TrimEnd('/') + "/components/prism-" + lang + ".min.js";
            if (mode == MarkdownAssetMode.Online) yield return new MarkdownHtmlAsset($"prism-lang:{lang}", MarkdownHtmlAssetKind.Js, url, null);
            else yield return new MarkdownHtmlAsset($"prism-lang:{lang}", MarkdownHtmlAssetKind.Js, null, HtmlRenderer.ResolveExternalText(url, resolver));
        }

        // Plugins (CSS+JS when available)
        foreach (var plugin in po.Plugins) {
            string js = po.CdnBase.TrimEnd('/') + "/plugins/" + plugin + "/prism-" + plugin + ".min.js";
            string css = po.CdnBase.TrimEnd('/') + "/plugins/" + plugin + "/prism-" + plugin + ".min.css";
            if (cssDelivery == MarkdownCssDelivery.LinkHref && mode == MarkdownAssetMode.Online) yield return new MarkdownHtmlAsset($"prism-plugin-css:{plugin}", MarkdownHtmlAssetKind.Css, css, null);
            else yield return new MarkdownHtmlAsset($"prism-plugin-css:{plugin}", MarkdownHtmlAssetKind.Css, null, HtmlRenderer.ScopeCss(HtmlRenderer.ResolveExternalText(css, resolver), scopeSelector));
            if (mode == MarkdownAssetMode.Online) yield return new MarkdownHtmlAsset($"prism-plugin:{plugin}", MarkdownHtmlAssetKind.Js, js, null);
            else yield return new MarkdownHtmlAsset($"prism-plugin:{plugin}", MarkdownHtmlAssetKind.Js, null, HtmlRenderer.ResolveExternalText(js, resolver));
        }
    }
}
