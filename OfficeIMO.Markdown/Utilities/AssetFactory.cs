namespace OfficeIMO.Markdown;

internal static class AssetFactory {
    internal static System.Collections.Generic.IEnumerable<HtmlAsset> PrismAssets(PrismOptions po, AssetMode mode, CssDelivery cssDelivery, string? scopeSelector) {
        // Build URLs
        string coreJs = po.CdnBase.TrimEnd('/') + "/components/prism-core.min.js";
        string prismLightCss = po.CdnBase.TrimEnd('/') + "/themes/prism.min.css";
        string prismDarkCss = po.CdnBase.TrimEnd('/') + "/themes/prism-okaidia.min.css";

        // Theme CSS handling
        if (po.Theme == PrismTheme.GithubAuto) {
            if (cssDelivery == CssDelivery.LinkHref && mode == AssetMode.Online) {
                var light = new HtmlAsset("prism-theme:light", HtmlAssetKind.Css, prismLightCss, null) { Media = "(prefers-color-scheme: light)" };
                var dark = new HtmlAsset("prism-theme:dark", HtmlAssetKind.Css, prismDarkCss, null) { Media = "(prefers-color-scheme: dark)" };
                yield return light; yield return dark;
            } else {
                var lightCss = HtmlRenderer.TryDownloadText(prismLightCss);
                var darkCss = HtmlRenderer.TryDownloadText(prismDarkCss);
                if (!string.IsNullOrEmpty(lightCss)) lightCss = HtmlRenderer.ScopeCss(lightCss, scopeSelector);
                if (!string.IsNullOrEmpty(darkCss)) darkCss = "@media (prefers-color-scheme: dark){\n" + HtmlRenderer.ScopeCss(darkCss, scopeSelector) + "\n}";
                yield return new HtmlAsset("prism-theme:auto", HtmlAssetKind.Css, null, (lightCss ?? string.Empty) + (string.IsNullOrEmpty(lightCss) || string.IsNullOrEmpty(darkCss) ? string.Empty : "\n") + (darkCss ?? string.Empty));
            }
        } else {
            string themeCss = po.Theme switch {
                PrismTheme.Okaidia => prismDarkCss,
                PrismTheme.GithubDark => prismDarkCss,
                _ => prismLightCss
            };
            if (cssDelivery == CssDelivery.LinkHref && mode == AssetMode.Online) {
                yield return new HtmlAsset($"prism-theme:{po.Theme}", HtmlAssetKind.Css, themeCss, null);
            } else {
                var css = HtmlRenderer.TryDownloadText(themeCss);
                if (!string.IsNullOrEmpty(css)) css = HtmlRenderer.ScopeCss(css!, scopeSelector);
                yield return new HtmlAsset($"prism-theme:{po.Theme}", HtmlAssetKind.Css, null, css);
            }
        }

        // Core JS
        if (mode == AssetMode.Online) yield return new HtmlAsset("prism-core", HtmlAssetKind.Js, coreJs, null);
        else yield return new HtmlAsset("prism-core", HtmlAssetKind.Js, null, HtmlRenderer.TryDownloadText(coreJs));

        // Languages
        foreach (var lang in po.Languages) {
            string url = po.CdnBase.TrimEnd('/') + "/components/prism-" + lang + ".min.js";
            if (mode == AssetMode.Online) yield return new HtmlAsset($"prism-lang:{lang}", HtmlAssetKind.Js, url, null);
            else yield return new HtmlAsset($"prism-lang:{lang}", HtmlAssetKind.Js, null, HtmlRenderer.TryDownloadText(url));
        }

        // Plugins (CSS+JS when available)
        foreach (var plugin in po.Plugins) {
            string js = po.CdnBase.TrimEnd('/') + "/plugins/" + plugin + "/prism-" + plugin + ".min.js";
            string css = po.CdnBase.TrimEnd('/') + "/plugins/" + plugin + "/prism-" + plugin + ".min.css";
            if (cssDelivery == CssDelivery.LinkHref && mode == AssetMode.Online) yield return new HtmlAsset($"prism-plugin-css:{plugin}", HtmlAssetKind.Css, css, null);
            else yield return new HtmlAsset($"prism-plugin-css:{plugin}", HtmlAssetKind.Css, null, HtmlRenderer.ScopeCss(HtmlRenderer.TryDownloadText(css), scopeSelector));
            if (mode == AssetMode.Online) yield return new HtmlAsset($"prism-plugin:{plugin}", HtmlAssetKind.Js, js, null);
            else yield return new HtmlAsset($"prism-plugin:{plugin}", HtmlAssetKind.Js, null, HtmlRenderer.TryDownloadText(js));
        }
    }
}

