using System.Net.Http;
using System.Threading;

namespace OfficeIMO.Markdown;

/// <summary>
/// Composes HTML fragments/documents from a MarkdownDoc with options.
/// </summary>
internal static class HtmlRenderer {
    internal static string Render(MarkdownDoc doc, HtmlOptions options) {
        var parts = RenderParts(doc, options);
        if (options.Kind == HtmlKind.Fragment) {
            return parts.Body; // Body already wrapped if requested
        }
        // Full document
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        sb.Append("<title>").Append(System.Net.WebUtility.HtmlEncode(options.Title ?? "Document")).Append("</title>");
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        if (!string.IsNullOrEmpty(parts.Head)) sb.Append(parts.Head);
        sb.Append("</head><body>");
        sb.Append(parts.Body);
        if (!string.IsNullOrEmpty(parts.Scripts)) sb.Append("<script>\n").Append(parts.Scripts).Append("\n</script>");
        sb.Append("</body></html>");
        return sb.ToString();
    }

    internal static HtmlRenderParts RenderParts(MarkdownDoc doc, HtmlOptions options) {
        using var _ctx = HtmlRenderContext.Push(options);
        var (realizedBlocks, headingCatalog) = doc.GetBlocksAndHeadingSlugs();
        var css = BuildCss(options, out string? cssLinkTag, out string? cssToWrite, out string? extraHeadLinks);
        options._externalCssContentToWrite = cssToWrite; // pass back for SaveHtml

        // Insert a top anchor for back-to-top links
        var blocksForRendering = doc.Blocks;
        string bodyContent = (options.BackToTopLinks ? "<a id=\"top\"></a>" : string.Empty) + RenderBody(blocksForRendering, options, headingCatalog);
        if (options.ThemeToggle) {
            const string toggle = "<button class=\"theme-toggle\" data-theme-toggle title=\"Toggle theme\" aria-label=\"Toggle theme\">🌓</button>";
            bodyContent = toggle + bodyContent;
        }
        if (!string.IsNullOrEmpty(options.BodyClass)) {
            // Wrap in article
            bodyContent = $"<article class=\"{System.Net.WebUtility.HtmlEncode(options.BodyClass)}\">" + bodyContent + "</article>";
        }

        StringBuilder head = new StringBuilder();
        if (!string.IsNullOrEmpty(cssLinkTag)) head.Append(cssLinkTag);
        if (!string.IsNullOrEmpty(extraHeadLinks)) head.Append(extraHeadLinks);

        StringBuilder scripts = new StringBuilder();
        if (options.ThemeToggle) scripts.Append(HtmlResources.ThemeToggleScript);
        if (options.CopyHeadingLinkOnClick) scripts.Append(HtmlResources.AnchorCopyScript);
        // ScrollSpy: include only if any TOC requests it
        try {
            if (doc.Blocks != null && doc.Blocks.Any(b => b is ITocPlaceholderMarkdownBlock toc && toc.RequiresScrollSpy())) {
                scripts.Append(HtmlResources.ScrollSpyScript);
            }
        } catch { /* best-effort */ }

        // Additional JS: link in head when Online; download+inline into scripts when Offline
        if (options.AssetMode == AssetMode.Online) {
            foreach (var js in options.AdditionalJsHrefs.Where(u => !string.IsNullOrWhiteSpace(u))) {
                head.Append($"<script src=\"{System.Net.WebUtility.HtmlEncode(js)}\"></script>\n");
            }
        } else {
            foreach (var js in options.AdditionalJsHrefs.Where(u => !string.IsNullOrWhiteSpace(u))) {
                var code = TryDownloadText(js);
                if (!string.IsNullOrEmpty(code)) scripts.Append(code).Append('\n');
            }
        }

        var parts = new HtmlRenderParts();

        // Prism assets (manifest + optional emission)
        if (options.Prism?.Enabled == true) {
            // For Prism in Online mode, prefer link-based CSS for GithubAuto theme to expose media attributes
            // so hosts can dedupe/merge correctly (tests expect media queries present in <link> tags).
            var prismCssDelivery = (options.AssetMode == AssetMode.Online && options.Prism.Theme == PrismTheme.GithubAuto)
                ? CssDelivery.LinkHref
                : options.CssDelivery;
            var assets = AssetFactory.PrismAssets(options.Prism, options.AssetMode, prismCssDelivery, options.CssScopeSelector);
            foreach (var a in assets) parts.Assets.Add(a);
            if (options.EmitMode == AssetEmitMode.Emit) {
                foreach (var a in parts.Assets) {
                    if (a.Kind == HtmlAssetKind.Css) {
                        if (!string.IsNullOrEmpty(a.Href)) {
                            var media = string.IsNullOrEmpty(a.Media) ? string.Empty : $" media=\"{System.Net.WebUtility.HtmlEncode(a.Media)}\"";
                            head.Append($"<link rel=\"stylesheet\" data-asset-id=\"{System.Net.WebUtility.HtmlEncode(a.Id)}\" href=\"{System.Net.WebUtility.HtmlEncode(a.Href)}\"{media}>\n");
                        } else if (!string.IsNullOrEmpty(a.Inline)) {
                            if (options.CssDelivery == CssDelivery.ExternalFile) {
                                // When writing CSS to a sidecar file, ensure Prism inline CSS is included there too.
                                options._externalCssContentToWrite = (options._externalCssContentToWrite?.Length > 0 ? (options._externalCssContentToWrite + "\n") : (options._externalCssContentToWrite ?? string.Empty)) + a.Inline;
                            } else {
                                css = (css?.Length > 0 ? (css + "\n") : (css ?? string.Empty)) + a.Inline;
                            }
                        }
                    } else {
                        if (!string.IsNullOrEmpty(a.Href)) head.Append($"<script data-asset-id=\"{System.Net.WebUtility.HtmlEncode(a.Id)}\" src=\"{System.Net.WebUtility.HtmlEncode(a.Href)}\"></script>\n");
                        else if (!string.IsNullOrEmpty(a.Inline)) scripts.Append(a.Inline).Append('\n');
                    }
                }
            }
        }

        // Capture final strings after optional asset emission.
        parts.Head = head.ToString();
        parts.Body = bodyContent;
        parts.Css = css ?? string.Empty;
        parts.Scripts = scripts.ToString();
        return parts;
    }

    private static string RenderBody(System.Collections.Generic.IReadOnlyList<IMarkdownBlock> blocks, HtmlOptions options, MarkdownHeadingCatalog headingCatalog) {
        var context = new MarkdownBodyRenderContext(blocks, options, headingCatalog);
        var plan = MarkdownBodyRenderPlan.Create(blocks);
        var footnotes = plan.Footnotes;
        var sidebar = plan.Sidebar;

        if (sidebar != null) {
            var navHtml = sidebar.RenderHtml(context);
            var content = new StringBuilder();
            for (int i = 0; i < plan.RenderBlocks.Count; i++) {
                content.Append(RenderBodyBlock(plan.RenderBlocks[i], context));
            }
        if (footnotes.Count > 0) content.Append(BuildFootnotesSectionHtml(footnotes, options));
        return sidebar.WrapSidebarLayoutHtml(navHtml, content.ToString());
    }

        var sb = new StringBuilder();
        for (int i = 0; i < plan.RenderBlocks.Count; i++) {
            sb.Append(RenderBodyBlock(plan.RenderBlocks[i], context));
        }
        if (footnotes.Count > 0) sb.Append(BuildFootnotesSectionHtml(footnotes, options));
        return sb.ToString();
    }

    private static string RenderBodyBlock(IMarkdownBlock block, MarkdownBodyRenderContext context) {
        var overridden = TryRenderBlockOverride(block, context.Options);
        if (overridden != null) {
            return overridden;
        }

        if (block is IContextualHtmlMarkdownBlock contextualBlock) {
            return contextualBlock.RenderHtml(context);
        }

        return block.RenderHtml();
    }

    private static string? TryRenderBlockOverride(IMarkdownBlock block, HtmlOptions options) {
        var extensions = options.BlockRenderExtensions;
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(block)) {
                continue;
            }

            var rendered = extension.RenderHtml(block, options);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string BuildFootnotesSectionHtml(IReadOnlyList<IFootnoteSectionMarkdownBlock> footnotes, HtmlOptions options) {
        if (footnotes == null || footnotes.Count == 0) return string.Empty;

        var typedFootnotes = footnotes.OfType<FootnoteDefinitionBlock>().ToList();
        var overridden = options.FootnoteSectionHtmlRenderer?.Invoke(typedFootnotes, options);
        if (overridden != null) {
            return overridden;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sb = new StringBuilder();
        sb.Append("<section class=\"footnotes\"><hr /><ol>");

        for (int i = 0; i < footnotes.Count; i++) {
            var fn = footnotes[i];
            if (fn == null) continue;

            var label = fn.FootnoteLabel ?? string.Empty;
            if (label.Length == 0) continue;
            if (!seen.Add(label)) continue;
            sb.Append(fn.RenderFootnoteSectionItemHtml());
        }

        sb.Append("</ol></section>");
        return sb.ToString();
    }

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _scopedBaseCssCache = new System.Collections.Concurrent.ConcurrentDictionary<string, string>(System.StringComparer.Ordinal);

    private static string? BuildCss(HtmlOptions options, out string? cssLinkTag, out string? cssToWrite, out string? extraHeadLinks) {
        cssLinkTag = null; cssToWrite = null; extraHeadLinks = null;
        // Cache scoped base CSS (style preset + common extras) by (style|scopeSelector)
        string cacheKey = ((int)options.Style).ToString() + "|" + (options.CssScopeSelector ?? string.Empty);
        if (!_scopedBaseCssCache.TryGetValue(cacheKey, out var baseCss)) {
            baseCss = ScopeCss(HtmlResources.GetStyleCss(options.Style) + HtmlResources.CommonExtraCss, options.CssScopeSelector);
            _scopedBaseCssCache[cacheKey] = baseCss;
        }

        // Additional CSS/JS URLs may be included in head as link/script or inlined depending on AssetMode
        StringBuilder headLinks = new StringBuilder();

        // Primary stylesheet selection
        if (options.CssDelivery == CssDelivery.None) {
            // Still emit links for additional CSS if Online
            if (options.AssetMode == AssetMode.Online) {
                foreach (var href in options.AdditionalCssHrefs.Where(u => !string.IsNullOrWhiteSpace(u)))
                    headLinks.Append($"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(href)}\">\n");
            }
            // AdditionalJs handled later in head (scripts in body for full doc)
            extraHeadLinks = headLinks.ToString();
            return string.Empty;
        }

        bool linkPrimary = false;
        if (options.CssDelivery == CssDelivery.LinkHref && !string.IsNullOrWhiteSpace(options.CssHref) && options.AssetMode == AssetMode.Online) {
            linkPrimary = true;
            cssLinkTag = $"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(options.CssHref)}\">\n";
            foreach (var href in options.AdditionalCssHrefs.Where(u => !string.IsNullOrWhiteSpace(u)))
                headLinks.Append($"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(href)}\">\n");
            extraHeadLinks = headLinks.ToString();
            // Do not return; we may inline small theme overrides below
        }

        // Inline or ExternalFile, or LinkHref with Offline mode
        var cssBuilder = new StringBuilder();
        if (!linkPrimary && !string.IsNullOrEmpty(baseCss)) cssBuilder.Append(baseCss).Append('\n');

        if (options.CssDelivery == CssDelivery.LinkHref && !string.IsNullOrWhiteSpace(options.CssHref) && options.AssetMode == AssetMode.Offline) {
            // Attempt to download provided CSS and inline
            var downloaded = TryDownloadText(options.CssHref!);
            if (!string.IsNullOrEmpty(downloaded)) cssBuilder.Append(downloaded).Append('\n');
        }
        // Additional CSS URLs
        foreach (var href in options.AdditionalCssHrefs.Where(u => !string.IsNullOrWhiteSpace(u))) {
            if (options.AssetMode == AssetMode.Online && options.CssDelivery == CssDelivery.LinkHref) {
                headLinks.Append($"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(href)}\">\n");
            } else {
                var downloaded = TryDownloadText(href);
                if (!string.IsNullOrEmpty(downloaded)) cssBuilder.Append(downloaded).Append('\n');
            }
        }
        extraHeadLinks = headLinks.ToString();

        // Theme overrides appended last so they win
        var overrides = BuildThemeOverrides(options);
        if (!string.IsNullOrEmpty(overrides)) cssBuilder.Append(overrides);
        var aggregatedCss = cssBuilder.ToString();
        if (options.CssDelivery == CssDelivery.ExternalFile) {
            // Renderer expects caller to write this CSS; return empty inline CSS but set writable content
            cssToWrite = aggregatedCss;
            var fileName = options.ExternalCssOutputPath != null ? System.IO.Path.GetFileName(options.ExternalCssOutputPath) : "styles.css";
            var styleId = $"omd-style:{options.Style}";
            cssLinkTag = $"<link rel=\"stylesheet\" data-asset-id=\"{System.Net.WebUtility.HtmlEncode(styleId)}\" href=\"{System.Net.WebUtility.HtmlEncode(fileName)}\">\n";
            return string.Empty;
        }
        return aggregatedCss;
    }

    private static string BuildThemeOverrides(HtmlOptions options) {
        var t = options.Theme ?? new ThemeColors();
        bool any = !string.IsNullOrWhiteSpace(t.AccentLight) || !string.IsNullOrWhiteSpace(t.AccentDark)
                 || !string.IsNullOrWhiteSpace(t.HeadingLight) || !string.IsNullOrWhiteSpace(t.HeadingDark)
                 || !string.IsNullOrWhiteSpace(t.TocBgLight) || !string.IsNullOrWhiteSpace(t.TocBgDark)
                 || !string.IsNullOrWhiteSpace(t.TocBorderLight) || !string.IsNullOrWhiteSpace(t.TocBorderDark)
                 || !string.IsNullOrWhiteSpace(t.ActiveLinkLight) || !string.IsNullOrWhiteSpace(t.ActiveLinkDark);
        if (!any) return string.Empty;
        var sb = new StringBuilder();
        var scope = NormalizeScope(options.CssScopeSelector);
        // Expose variables on scope for both themes
        sb.Append('\n');
        sb.Append(CombineSelectors("html[data-theme=light]", scope)).Append(" {");
        if (!string.IsNullOrWhiteSpace(t.HeadingLight)) sb.Append(" --md-heading: ").Append(t.HeadingLight).Append(';');
        if (!string.IsNullOrWhiteSpace(t.AccentLight)) sb.Append(" --md-accent: ").Append(t.AccentLight).Append(';');
        if (!string.IsNullOrWhiteSpace(t.TocBgLight)) sb.Append(" --md-toc-bg: ").Append(t.TocBgLight).Append(';');
        if (!string.IsNullOrWhiteSpace(t.TocBorderLight)) sb.Append(" --md-toc-border: ").Append(t.TocBorderLight).Append(';');
        if (!string.IsNullOrWhiteSpace(t.ActiveLinkLight)) sb.Append(" --md-active: ").Append(t.ActiveLinkLight).Append(';');
        sb.Append(" }\n");
        sb.Append(CombineSelectors("html[data-theme=dark]", scope)).Append(" {");
        if (!string.IsNullOrWhiteSpace(t.HeadingDark)) sb.Append(" --md-heading: ").Append(t.HeadingDark).Append(';');
        if (!string.IsNullOrWhiteSpace(t.AccentDark)) sb.Append(" --md-accent: ").Append(t.AccentDark).Append(';');
        if (!string.IsNullOrWhiteSpace(t.TocBgDark)) sb.Append(" --md-toc-bg: ").Append(t.TocBgDark).Append(';');
        if (!string.IsNullOrWhiteSpace(t.TocBorderDark)) sb.Append(" --md-toc-border: ").Append(t.TocBorderDark).Append(';');
        if (!string.IsNullOrWhiteSpace(t.ActiveLinkDark)) sb.Append(" --md-active: ").Append(t.ActiveLinkDark).Append(';');
        sb.Append(" }\n");
        // Map variables to elements
        sb.Append(string.Join(", ", Descendant(scope, "h1"), Descendant(scope, "h2"), Descendant(scope, "h3"),
            Descendant(scope, "h4"), Descendant(scope, "h5"), Descendant(scope, "h6")))
          .Append(" { color: var(--md-heading, inherit); }\n");
        sb.Append(Descendant(scope, "a")).Append(" { color: var(--md-accent, #0969da); }\n");
        sb.Append(Descendant(scope, "nav.md-toc")).Append(" { background: var(--md-toc-bg, #f6f8fa); border-color: var(--md-toc-border, #d0d7de); }\n");
        sb.Append(Descendant(scope, "nav.md-toc a.active")).Append(" { color: var(--md-active, var(--md-accent, #0969da)); border-left-color: var(--md-active, var(--md-accent, #0969da)); }\n");
        // Accented borders and anchors for stronger theme feel
        sb.Append(Descendant(scope, "h2")).Append(" { border-bottom-color: var(--md-accent, #d8dee4); }\n");
        sb.Append(Descendant(scope, "blockquote")).Append(" { border-left-color: var(--md-accent, #d0d7de); }\n");
        sb.Append(Descendant(scope, "blockquote.callout")).Append(" { border-left-color: var(--md-accent, #0969da); background: var(--md-toc-bg, #f6f8fa); }\n");
        sb.Append(Descendant(scope, ".heading-anchor")).Append(" { color: var(--md-accent, inherit); }\n");
        return sb.ToString();
    }

    internal static string TryDownloadText(string? url) {
        try {
            if (string.IsNullOrWhiteSpace(url)) return string.Empty;
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return string.Empty;
            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) return string.Empty;

            const long MaxBytes = 1_000_000; // 1MB guardrail
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(8));
            var handler = new HttpClientHandler { AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate };
            using var client = new HttpClient(handler);
            using var resp = client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cts.Token).GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode) return string.Empty;
            var ct = resp.Content.Headers.ContentType?.MediaType?.ToLowerInvariant();
            bool okType = false;
            if (ct == null) okType = true;
            else if (ct.StartsWith("text/")) okType = true;
            else if (ct.IndexOf("javascript", StringComparison.Ordinal) >= 0 || ct.IndexOf("ecmascript", StringComparison.Ordinal) >= 0 || ct.IndexOf("css", StringComparison.Ordinal) >= 0) okType = true;
            if (!okType) return string.Empty;
            var len = resp.Content.Headers.ContentLength;
            if (len.HasValue && len.Value > MaxBytes) return string.Empty;
            using var stream = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
            using var mem = new System.IO.MemoryStream(len.HasValue ? (int)Math.Min(len.Value, MaxBytes) : 64 * 1024);
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                total += read;
                if (total > MaxBytes) return string.Empty;
                mem.Write(buffer, 0, read);
            }
            var charset = resp.Content.Headers.ContentType?.CharSet;
            Encoding enc;
            try { enc = !string.IsNullOrWhiteSpace(charset) ? Encoding.GetEncoding(charset!) : new UTF8Encoding(false); } catch { enc = new UTF8Encoding(false); }
            return enc.GetString(mem.ToArray());
        } catch { return string.Empty; }
    }

    internal static string ScopeCss(string? css, string? scopeSelector) {
        string cssText = css ?? string.Empty;
        if (cssText.Length == 0) return string.Empty;
        string scope = NormalizeScope(scopeSelector);
        // Naive scoping: prefix common selectors with the scope to avoid global bleed.
        // This is intentionally conservative.
        var s = cssText.Replace("code[class*=\"language-\"]", scope + " code[class*=\\\"language-\\\"]")
                   .Replace("pre[class*=\"language-\"]", scope + " pre[class*=\\\"language-\\\"]")
                   .Replace("pre[class*=\"language-\"] code", scope + " pre[class*=\\\"language-\\\"] code");
        // Also prefix top-level element rules we own
        s = s.Replace("article.markdown-body", scope);
        return s;
    }

    private static string NormalizeScope(string? scopeSelector) {
        string selector = scopeSelector ?? string.Empty;
        if (string.IsNullOrWhiteSpace(selector)) return "body";
        return selector.Trim();
    }

    private static string CombineSelectors(string prefix, string scope) {
        if (string.IsNullOrEmpty(prefix)) return scope;
        if (string.IsNullOrEmpty(scope)) return prefix;
        return prefix + " " + scope;
    }

    private static string Descendant(string scope, string selector) {
        if (string.IsNullOrEmpty(scope)) return selector;
        if (string.IsNullOrEmpty(selector)) return scope;
        return scope + " " + selector;
    }
}
