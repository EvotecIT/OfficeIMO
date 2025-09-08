using System;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;

namespace OfficeIMO.Markdown.Utilities;

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
        var realized = GetBlocks(doc);
        var css = BuildCss(options, out string? cssLinkTag, out string? cssToWrite, out string? extraHeadLinks);
        options._externalCssContentToWrite = cssToWrite; // pass back for SaveHtml

        // Insert a top anchor for back-to-top links
        string bodyContent = (options.BackToTopLinks ? "<a id=\"top\"></a>" : string.Empty) + RenderBody(realized, options);
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

        var parts = new HtmlRenderParts {
            Head = head.ToString(),
            Body = bodyContent,
            Css = css ?? string.Empty,
            Scripts = scripts.ToString()
        };

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
                        }
                        else if (!string.IsNullOrEmpty(a.Inline)) parts.Css += (parts.Css.Length > 0 ? "\n" : "") + a.Inline;
                    } else {
                        if (!string.IsNullOrEmpty(a.Href)) head.Append($"<script data-asset-id=\"{System.Net.WebUtility.HtmlEncode(a.Id)}\" src=\"{System.Net.WebUtility.HtmlEncode(a.Href)}\"></script>\n");
                        else if (!string.IsNullOrEmpty(a.Inline)) parts.Scripts += (parts.Scripts.Length > 0 ? "\n" : "") + a.Inline;
                    }
                }
            }
        }

        return parts;
    }

    private static System.Collections.Generic.List<IMarkdownBlock> GetBlocks(MarkdownDoc doc) {
        // Access private RealizeTocPlaceholders via public ToMarkdown? We replicate the realization by re-rendering via Markdown
        // but that would lose block info. Instead, since RealizeTocPlaceholders is private, we emulate quickly:
        // Render Markdown and re-parse? Too heavy. Simpler approach: copy internal logic here by reflection.
        // As we control code, we call doc.ToMarkdown() to force TOC realization, but we still have _blocks.
        // We will re-run generation by temporarily creating a new doc from markdown is not feasible.
        // Instead: use the HTML from doc.ToHtml() for non-heading blocks and stitch headings with anchors.
        // For now, rely on doc.ToHtml() which already realizes TOC. This returns concatenated block HTML with ids.
        return new System.Collections.Generic.List<IMarkdownBlock>(doc.Blocks);
    }

    private static string RenderBody(System.Collections.Generic.IReadOnlyList<IMarkdownBlock> blocks, HtmlOptions options) {
        var sb = new StringBuilder();
        foreach (var block in blocks) {
            if (block is HeadingBlock h) {
                var id = MarkdownSlug.GitHub(h.Text);
                var encoded = System.Net.WebUtility.HtmlEncode(h.Text);
                sb.Append($"<h{h.Level} id=\"{id}\">");
                sb.Append(encoded);
                if (options.IncludeAnchorLinks || options.ShowAnchorIcons) {
                    var icon = System.Net.WebUtility.HtmlEncode(options.AnchorIcon ?? "🔗");
                    // Anchor icon button; when CopyHeadingLinkOnClick, JS hooks it to copy full URL
                    sb.Append($"<a class=\"heading-anchor\" href=\"#{id}\" data-anchor-id=\"{id}\" title=\"Copy link\" aria-label=\"Copy link\">{icon}</a>");
                }
                sb.Append($"</h{h.Level}>");
                if (options.BackToTopLinks && h.Level >= options.BackToTopMinLevel) {
                    var txt = System.Net.WebUtility.HtmlEncode(options.BackToTopText ?? "Back to top");
                    sb.Append($"<div class=\"back-to-top\"><a href=\"#top\">{txt}</a></div>");
                }
            } else if (block is TocPlaceholderBlock tp) {
                sb.Append(BuildTocHtml(blocks, tp));
            } else {
                sb.Append(block.RenderHtml());
            }
        }
        return sb.ToString();
    }

    private static string BuildTocHtml(System.Collections.Generic.IReadOnlyList<IMarkdownBlock> blocks, TocPlaceholderBlock tp) {
        var opts = tp.Options;
        // Collect headings with indices
        var headings = blocks
            .Select((b, idx) => (b, idx))
            .Where(t => t.b is HeadingBlock)
            .Select(t => (Index: t.idx, Level: ((HeadingBlock)t.b).Level, Text: ((HeadingBlock)t.b).Text))
            .ToList();

        int placeholderIndex = System.Array.IndexOf(blocks.ToArray(), tp);
        int startIdx = 0; int endIdx = blocks.Count;
        if (opts.Scope == TocScope.PreviousHeading) {
            var prev = headings.LastOrDefault(h => h.Index < placeholderIndex && h.Level < opts.MinLevel);
            if (prev.Equals(default((int,int,string)))) prev = headings.LastOrDefault(h => h.Index < placeholderIndex);
            if (!prev.Equals(default((int,int,string)))) {
                startIdx = prev.Index + 1;
                var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > prev.Index && h.Level <= prev.Level);
                if (!nextAtOrAbove.Equals(default((int,int,string)))) endIdx = nextAtOrAbove.Index;
            }
        } else if (opts.Scope == TocScope.HeadingTitle && !string.IsNullOrWhiteSpace(opts.ScopeHeadingTitle)) {
            var root = headings.FirstOrDefault(h => string.Equals(h.Text, opts.ScopeHeadingTitle, StringComparison.OrdinalIgnoreCase));
            if (!root.Equals(default((int,int,string)))) {
                startIdx = root.Index + 1;
                var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > root.Index && h.Level <= root.Level);
                if (!nextAtOrAbove.Equals(default((int,int,string)))) endIdx = nextAtOrAbove.Index;
            }
        }

        var relevant = headings.Where(h => h.Index >= startIdx && h.Index < endIdx && h.Level >= opts.MinLevel && h.Level <= opts.MaxLevel)
                               .Select(h => (h.Level, h.Text, Anchor: MarkdownSlug.GitHub(h.Text)))
                               .ToList();
        if (opts.IncludeTitle && !string.IsNullOrWhiteSpace(opts.Title)) {
            var titleSlug = MarkdownSlug.GitHub(opts.Title);
            relevant = relevant.Where(e => !string.Equals(e.Anchor, titleSlug, StringComparison.Ordinal)).ToList();
        }
        if (relevant.Count == 0) return string.Empty;

        // Build nested list of headings
        var listTag = opts.Ordered ? "ol" : "ul";
        var sbNested = new StringBuilder();
        int baseLevel = relevant.Min(r => r.Level);
        int current = baseLevel - 1;
        while (current < baseLevel) { sbNested.Append('<').Append(listTag).Append('>'); current++; }
        foreach (var e in relevant) {
            while (current < e.Level) { sbNested.Append('<').Append(listTag).Append('>'); current++; }
            while (current > e.Level) { sbNested.Append("</").Append(listTag).Append('>'); current--; }
            sbNested.Append("<li><a href=\"")
                    .Append('#').Append(System.Net.WebUtility.HtmlEncode(e.Anchor))
                    .Append("\">").Append(System.Net.WebUtility.HtmlEncode(e.Text))
                    .Append("</a></li>");
        }
        while (current >= baseLevel) { sbNested.Append("</").Append(listTag).Append('>'); current--; }

        if (opts.Collapsible) {
            string open = opts.Collapsed ? string.Empty : " open";
            string summary = System.Net.WebUtility.HtmlEncode(opts.Title ?? "Contents");
            var sbWrap = new StringBuilder();
            sbWrap.Append("<details class=\"md-toc\"").Append(open).Append("><summary>")
                  .Append(summary).Append("</summary>")
                  .Append(sbNested.ToString())
                  .Append("</details>");
            return sbWrap.ToString();
        }

        if (opts.IncludeTitle) {
            var sbo = new StringBuilder();
            sbo.Append("<h").Append(opts.TitleLevel).Append('>')
               .Append(System.Net.WebUtility.HtmlEncode(opts.Title))
               .Append("</h").Append(opts.TitleLevel).Append('>')
               .Append(sbNested.ToString());
            return sbo.ToString();
        }
        return sbNested.ToString();
    }

    private static string? BuildCss(HtmlOptions options, out string? cssLinkTag, out string? cssToWrite, out string? extraHeadLinks) {
        cssLinkTag = null; cssToWrite = null; extraHeadLinks = null;
        var baseCss = ScopeCss(HtmlResources.GetStyleCss(options.Style), options.CssScopeSelector);

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

        if (options.CssDelivery == CssDelivery.LinkHref && !string.IsNullOrWhiteSpace(options.CssHref) && options.AssetMode == AssetMode.Online) {
            cssLinkTag = $"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(options.CssHref)}\">\n";
            foreach (var href in options.AdditionalCssHrefs.Where(u => !string.IsNullOrWhiteSpace(u)))
                headLinks.Append($"<link rel=\"stylesheet\" href=\"{System.Net.WebUtility.HtmlEncode(href)}\">\n");
            extraHeadLinks = headLinks.ToString();
            return string.Empty; // No inline CSS, referenced via link
        }

        // Inline or ExternalFile, or LinkHref with Offline mode
        var cssBuilder = new StringBuilder();
        if (!string.IsNullOrEmpty(baseCss)) cssBuilder.Append(baseCss).Append('\n');

        if (options.CssDelivery == CssDelivery.LinkHref && !string.IsNullOrWhiteSpace(options.CssHref) && options.AssetMode == AssetMode.Offline) {
            // Attempt to download provided CSS and inline
            var downloaded = TryDownloadText(options.CssHref);
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

    internal static string TryDownloadText(string url) {
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
            // Allow common text assets used by this library
            bool okType = false;
            if (ct == null) okType = true; // be permissive when unknown
            else if (ct.StartsWith("text/")) okType = true;
            else if (ct.Contains("javascript") || ct.Contains("ecmascript") || ct.Contains("css")) okType = true;
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
            // Determine encoding
            var charset = resp.Content.Headers.ContentType?.CharSet;
            Encoding enc;
            try { enc = !string.IsNullOrWhiteSpace(charset) ? Encoding.GetEncoding(charset!) : new UTF8Encoding(false); }
            catch { enc = new UTF8Encoding(false); }
            return enc.GetString(mem.ToArray());
        } catch { return string.Empty; }
    }

    internal static string ScopeCss(string? css, string scopeSelector) {
        if (string.IsNullOrEmpty(css)) return string.Empty;
        // Naive scoping: prefix common selectors with the scope to avoid global bleed.
        // This is intentionally conservative.
        var s = css.Replace("code[class*=\"language-\"]", scopeSelector + " code[class*=\\\"language-\\\"]")
                   .Replace("pre[class*=\"language-\"]", scopeSelector + " pre[class*=\\\"language-\\\"]")
                   .Replace("pre[class*=\"language-\"] code", scopeSelector + " pre[class*=\\\"language-\\\"] code");
        // Also prefix top-level element rules we own
        s = s.Replace("article.markdown-body", scopeSelector);
        return s;
    }
}
