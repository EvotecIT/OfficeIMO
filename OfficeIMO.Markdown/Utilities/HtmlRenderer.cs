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
        var (_, headingSlugs) = doc.GetBlocksAndHeadingSlugs();
        var css = BuildCss(options, out string? cssLinkTag, out string? cssToWrite, out string? extraHeadLinks);
        options._externalCssContentToWrite = cssToWrite; // pass back for SaveHtml

        // Insert a top anchor for back-to-top links
        var blocksForRendering = doc.Blocks;
        string bodyContent = (options.BackToTopLinks ? "<a id=\"top\"></a>" : string.Empty) + RenderBody(blocksForRendering, options, headingSlugs);
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
            if (doc.Blocks != null && doc.Blocks.Any(b => b is TocPlaceholderBlock tp && tp.Options.ScrollSpy)) {
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

    private static string RenderBody(System.Collections.Generic.IReadOnlyList<IMarkdownBlock> blocks, HtmlOptions options, System.Collections.Generic.IReadOnlyDictionary<HeadingBlock, string> headingSlugs) {
        // Footnote definitions are rendered at the bottom in a dedicated section.
        var footnotes = new List<FootnoteDefinitionBlock>();

        string RenderBlockHtml(IMarkdownBlock block) {
            if (block is HtmlRawBlock raw) {
                return options.RawHtmlHandling switch {
                    RawHtmlHandling.Allow => raw.Html,
                    RawHtmlHandling.Escape => "<pre class=\"md-raw-html\"><code>" + System.Net.WebUtility.HtmlEncode(raw.Html) + "</code></pre>",
                    _ => string.Empty
                };
            }
            if (block is HtmlCommentBlock c) {
                return options.RawHtmlHandling switch {
                    RawHtmlHandling.Allow => c.Comment,
                    RawHtmlHandling.Escape => "<pre class=\"md-raw-html\"><code>" + System.Net.WebUtility.HtmlEncode(c.Comment) + "</code></pre>",
                    _ => string.Empty
                };
            }
            return block.RenderHtml();
        }

        // Detect a sidebar TOC and render a two-column layout when present
        TocPlaceholderBlock? sidebar = null;
        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] is TocPlaceholderBlock tp && (tp.Options.Layout == TocLayout.SidebarLeft || tp.Options.Layout == TocLayout.SidebarRight)) {
                sidebar = tp; break;
            }
        }

        if (sidebar != null) {
            var side = sidebar.Options.Layout == TocLayout.SidebarLeft ? "left" : "right";
            var navHtml = BuildTocHtml(blocks, sidebar, headingSlugs);
            var content = new StringBuilder();
            for (int i = 0; i < blocks.Count; i++) {
                var block = blocks[i];
                if (block is FootnoteDefinitionBlock fn) { footnotes.Add(fn); continue; }
                // Skip the TOC title heading for enhanced layouts to avoid duplicate "On this page" in content
                if (block is HeadingBlock h0 && i + 1 < blocks.Count && blocks[i + 1] is TocPlaceholderBlock tp0) {
                    var o0 = tp0.Options;
                    if (o0.IncludeTitle && (o0.Layout == TocLayout.SidebarLeft || o0.Layout == TocLayout.SidebarRight || o0.Layout == TocLayout.Panel)) {
                        continue; // do not render this heading in content
                    }
                }
                if (ReferenceEquals(block, sidebar)) continue; // skip the sidebar placeholder; it's rendered as navHtml
                if (block is HeadingBlock h) {
                    if (!headingSlugs.TryGetValue(h, out var id)) id = MarkdownSlug.GitHub(h.Text);
                    var encoded = System.Net.WebUtility.HtmlEncode(h.Text);
                    content.Append($"<h{h.Level} id=\"{id}\">");
                    content.Append(encoded);
                    if (options.IncludeAnchorLinks || options.ShowAnchorIcons) {
                        var icon = System.Net.WebUtility.HtmlEncode(options.AnchorIcon ?? "🔗");
                        content.Append($"<a class=\"heading-anchor\" href=\"#{id}\" data-anchor-id=\"{id}\" title=\"Copy link\" aria-label=\"Copy link\">{icon}</a>");
                    }
                    content.Append($"</h{h.Level}>");
                    if (options.BackToTopLinks && h.Level >= options.BackToTopMinLevel) {
                        var txt = System.Net.WebUtility.HtmlEncode(options.BackToTopText ?? "Back to top");
                        content.Append($"<div class=\"back-to-top\"><a href=\"#top\">{txt}</a></div>");
                    }
                } else if (block is TocPlaceholderBlock tp) {
                    // Render any non-sidebar TOCs inline within content
                    if (!(tp.Options.Layout == TocLayout.SidebarLeft || tp.Options.Layout == TocLayout.SidebarRight))
                        content.Append(BuildTocHtml(blocks, tp, headingSlugs));
                } else {
                    content.Append(RenderBlockHtml(block));
                }
            }
            if (footnotes.Count > 0) content.Append(BuildFootnotesSectionHtml(footnotes));
            var sbLayout = new StringBuilder();
            string widthStyle = sidebar.Options.WidthPx.HasValue ? $" style=\"--md-toc-width: {sidebar.Options.WidthPx.Value}px\"" : string.Empty;
            sbLayout.Append($"<div class=\"md-layout two-col {side}\"{widthStyle}>");
            if (side == "left") {
                sbLayout.Append(navHtml).Append("<div class=\"md-content\">").Append(content.ToString()).Append("</div>");
            } else {
                sbLayout.Append("<div class=\"md-content\">").Append(content.ToString()).Append("</div>").Append(navHtml);
            }
            sbLayout.Append("</div>");
            return sbLayout.ToString();
        }

        var sb = new StringBuilder();
        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block is FootnoteDefinitionBlock fn) { footnotes.Add(fn); continue; }
            // Skip TOC title heading for enhanced layouts
            if (block is HeadingBlock h0 && i + 1 < blocks.Count && blocks[i + 1] is TocPlaceholderBlock tp0) {
                var o0 = tp0.Options;
                if (o0.IncludeTitle && (o0.Layout == TocLayout.SidebarLeft || o0.Layout == TocLayout.SidebarRight || o0.Layout == TocLayout.Panel)) {
                    continue;
                }
            }
            if (block is HeadingBlock h) {
                if (!headingSlugs.TryGetValue(h, out var id)) id = MarkdownSlug.GitHub(h.Text);
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
                sb.Append(BuildTocHtml(blocks, tp, headingSlugs));
            } else {
                sb.Append(RenderBlockHtml(block));
            }
        }
        if (footnotes.Count > 0) sb.Append(BuildFootnotesSectionHtml(footnotes));
        return sb.ToString();
    }

    private static string BuildFootnotesSectionHtml(List<FootnoteDefinitionBlock> footnotes) {
        if (footnotes == null || footnotes.Count == 0) return string.Empty;

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sb = new StringBuilder();
        sb.Append("<section class=\"footnotes\"><hr /><ol>");

        for (int i = 0; i < footnotes.Count; i++) {
            var fn = footnotes[i];
            if (fn == null) continue;

            var label = fn.Label ?? string.Empty;
            if (label.Length == 0) continue;
            if (!seen.Add(label)) continue;

            var enc = System.Net.WebUtility.HtmlEncode(label);
            sb.Append("<li id=\"fn:").Append(enc).Append("\">");

            var paragraphs = fn.Paragraphs;
            if (paragraphs == null || paragraphs.Count == 0) {
                var one = MarkdownReader.ParseInlineText(fn.Text);
                paragraphs = new List<InlineSequence> { one };
            }

            for (int p = 0; p < paragraphs.Count; p++) {
                var para = paragraphs[p] ?? new InlineSequence();
                sb.Append("<p>").Append(para.RenderHtml());
                if (p == paragraphs.Count - 1) {
                    sb.Append(" <a class=\"footnote-backref\" href=\"#fnref:").Append(enc).Append("\" aria-label=\"Back to reference\">&#8617;</a>");
                }
                sb.Append("</p>");
            }

            sb.Append("</li>");
        }

        sb.Append("</ol></section>");
        return sb.ToString();
    }

    private static string BuildTocHtml(System.Collections.Generic.IReadOnlyList<IMarkdownBlock> blocks, TocPlaceholderBlock tp, System.Collections.Generic.IReadOnlyDictionary<HeadingBlock, string> headingSlugs) {
        static int ClampHeadingLevel(int level) => level < 1 ? 1 : (level > 6 ? 6 : level);

        var opts = tp.Options;
        int minLevel = ClampHeadingLevel(opts.MinLevel);
        int maxLevel = ClampHeadingLevel(opts.MaxLevel);
        if (maxLevel < minLevel) maxLevel = minLevel;
        int titleLevel = ClampHeadingLevel(opts.TitleLevel);

        int effectiveMin = opts.RequireTopLevel && minLevel > 1 ? 1 : minLevel;
        int effectiveMax = maxLevel;
        // Collect headings with indices
        var headings = blocks
            .Select((b, idx) => (Index: idx, Heading: b as HeadingBlock))
            .Where(t => t.Heading is not null)
            .Select(t => (Index: t.Index, Level: t.Heading!.Level, Text: t.Heading.Text, Heading: t.Heading))
            .ToList();

        int placeholderIndex = System.Array.IndexOf(blocks.ToArray(), tp);
        int startIdx = 0; int endIdx = blocks.Count;
        if (opts.Scope == TocScope.PreviousHeading) {
            var prev = headings.LastOrDefault(h => h.Index < placeholderIndex && h.Level < minLevel);
            if (prev.Equals(default((int, int, string, HeadingBlock)))) prev = headings.LastOrDefault(h => h.Index < placeholderIndex);
            if (!prev.Equals(default((int, int, string, HeadingBlock)))) {
                startIdx = prev.Index + 1;
                var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > prev.Index && h.Level <= prev.Level);
                if (!nextAtOrAbove.Equals(default((int, int, string, HeadingBlock)))) endIdx = nextAtOrAbove.Index;
            }
        } else if (opts.Scope == TocScope.HeadingTitle && !string.IsNullOrWhiteSpace(opts.ScopeHeadingTitle)) {
            var root = headings.FirstOrDefault(h => string.Equals(h.Text, opts.ScopeHeadingTitle, StringComparison.OrdinalIgnoreCase));
            if (!root.Equals(default((int, int, string, HeadingBlock)))) {
                startIdx = root.Index + 1;
                var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > root.Index && h.Level <= root.Level);
                if (!nextAtOrAbove.Equals(default((int, int, string, HeadingBlock)))) endIdx = nextAtOrAbove.Index;
            }
        }

        var relevant = headings.Where(h => h.Index >= startIdx && h.Index < endIdx && h.Level >= effectiveMin && h.Level <= effectiveMax)
                               .Select(h => (h.Level, h.Text, Anchor: headingSlugs.TryGetValue(h.Heading, out var slug) ? slug : MarkdownSlug.GitHub(h.Text)))
                               .ToList();
        if (opts.IncludeTitle && !string.IsNullOrWhiteSpace(opts.Title) && placeholderIndex > 0 && blocks[placeholderIndex - 1] is HeadingBlock titleHeading) {
            if (!headingSlugs.TryGetValue(titleHeading, out var titleSlug)) titleSlug = MarkdownSlug.GitHub(titleHeading.Text);
            relevant = relevant.Where(e => !string.Equals(e.Anchor, titleSlug, StringComparison.Ordinal)).ToList();
        }
        if (relevant.Count == 0) return string.Empty;

        // Build nested list of headings, ensuring nested lists are children of their parent <li>
        // HTML shape produced (example):
        // <ul>
        //   <li>H2
        //     <ul>
        //       <li>H3</li>
        //       <li>H3</li>
        //     </ul>
        //   </li>
        //   <li>H2</li>
        // </ul>
        var listTag = opts.Ordered ? "ol" : "ul";
        var sbNested = new StringBuilder(relevant.Count * 64);
        int baseLevel = (opts.NormalizeToMinLevel ? relevant.Min(r => r.Level) : 1);
        int currentDepth = 0; // relative to baseLevel
        bool any = false;
        sbNested.Append('<').Append(listTag).Append('>');
        for (int i = 0; i < relevant.Count; i++) {
            var e = relevant[i];
            int depth = Math.Max(0, e.Level - baseLevel);

            if (any) {
                if (depth == currentDepth) {
                    // sibling at same depth: close previous item
                    sbNested.Append("</li>");
                } else if (depth > currentDepth) {
                    // dive: open nested lists inside the current <li>
                    for (int d = currentDepth; d < depth; d++) sbNested.Append('<').Append(listTag).Append('>');
                } else /* depth < currentDepth */ {
                    // ascend: close open item and unwind lists
                    for (int d = currentDepth; d > depth; d--) sbNested.Append("</li></").Append(listTag).Append('>');
                    // close the parent item at the new depth as we move to a sibling
                    sbNested.Append("</li>");
                }
            }

            sbNested.Append("<li><a href=\"")
                    .Append('#').Append(System.Net.WebUtility.HtmlEncode(e.Anchor))
                    .Append("\">")
                    .Append(System.Net.WebUtility.HtmlEncode(e.Text))
                    .Append("</a>");

            currentDepth = depth; any = true;
        }
        if (any) sbNested.Append("</li>");
        for (int d = currentDepth; d > 0; d--) sbNested.Append("</").Append(listTag).Append("></li>");
        sbNested.Append("</").Append(listTag).Append('>');

        // Collapsible panel (legacy + styled)
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

        // Enhanced layouts: wrap in <nav> with classes for styling/scrollspy when requested
        bool enhanced = opts.Layout != TocLayout.List || opts.ScrollSpy || opts.Sticky;
        if (enhanced) {
            var classes = new StringBuilder("md-toc");
            if (opts.Layout == TocLayout.Panel) classes.Append(" panel");
            if (opts.Layout == TocLayout.SidebarRight) classes.Append(" sidebar right");
            if (opts.Layout == TocLayout.SidebarLeft) classes.Append(" sidebar left");
            if (opts.Sticky) classes.Append(" sticky");
            if (opts.ScrollSpy) classes.Append(" md-scrollspy autoscroll");
            if (opts.Chrome == TocChrome.None) classes.Append(" no-chrome");
            if (opts.Chrome == TocChrome.Outline) classes.Append(" outline");
            if (opts.Chrome == TocChrome.Panel) classes.Append(" panel");
            if (opts.HideOnNarrow) classes.Append(" hide-narrow");
            string aria = System.Net.WebUtility.HtmlEncode(opts.AriaLabel ?? "Table of Contents");
            var sb = new StringBuilder();
            sb.Append("<nav role=\"navigation\" aria-label=\"").Append(aria).Append("\" class=\"").Append(classes).Append("\"");
            if (opts.ScrollSpy) sb.Append(" data-md-scrollspy=\"1\"");
            if (opts.Sticky) sb.Append(" data-autoscroll=\"1\"");
            sb.Append(">");
            if (opts.IncludeTitle && !string.IsNullOrWhiteSpace(opts.Title)) {
                sb.Append("<div class=\"toc-title\">").Append(System.Net.WebUtility.HtmlEncode(opts.Title)).Append("</div>");
            }
            sb.Append(sbNested.ToString());
            sb.Append("</nav>");
            return sb.ToString();
        }

        // Legacy: emit plain heading + list without wrapper
        if (opts.IncludeTitle) {
            var sbo = new StringBuilder();
            sbo.Append("<h").Append(titleLevel).Append('>')
               .Append(System.Net.WebUtility.HtmlEncode(opts.Title))
               .Append("</h").Append(titleLevel).Append('>')
               .Append(sbNested.ToString());
            return sbo.ToString();
        }
        return sbNested.ToString();
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
