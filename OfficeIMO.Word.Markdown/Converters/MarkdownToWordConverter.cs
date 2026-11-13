using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Markdown → Word converter powered by OfficeIMO.Markdown.
    /// Maps OMD blocks/inlines onto OfficeIMO.Word APIs (headings, lists, tables, images,
    /// code, quotes, callouts, footnotes, etc.).
    /// </summary>
    internal partial class MarkdownToWordConverter {
        private const int IndentTwipsPerLevel = 720; // 0.5 inch per level
        private const double DefaultPageWidthTwips = 12240d;
        private const double DefaultHorizontalMarginTwips = 1440d;
        private const double TwipsPerInch = 1440d;
        private const double PixelsPerInch = 96d;
        private const double MinimumContentWidthPixels = 1d;
        // Current footnote definitions map; scoped to this per-conversion converter instance.
        private IReadOnlyDictionary<string, string>? _currentFootnotes;

        public WordDocument Convert(Omd.MarkdownDocument markdown, MarkdownWordExportOptions options, CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(markdown);

            options ??= new MarkdownWordExportOptions();

            cancellationToken.ThrowIfCancellationRequested();
            OfficeTheme? sourceTheme = options.ThemeSnapshot?.EffectiveSourceTheme;
            // Preserve the adapter's established implicit heading appearance through the native owner.
            OfficeTheme theme = sourceTheme ?? OfficeTheme.Default.WithTextStyle(
                OfficeTheme.Default.TextStyle.With(headingColor: WordThemeDefaults.PageText(OfficeTheme.Default),
                    headingBold: true, lineSpacing: WordThemeDefaults.NativeBodyLineSpacing));
            var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
            try {
                if (sourceTheme != null) {
                    document.PageSetup.Margins.Top = sourceTheme.Page.MarginTop;
                    document.PageSetup.Margins.Right = sourceTheme.Page.MarginRight;
                    document.PageSetup.Margins.Bottom = sourceTheme.Page.MarginBottom;
                    document.PageSetup.Margins.Left = sourceTheme.Page.MarginLeft;
                }
                options.ApplyDefaults(document);
                var pageContentWidthPixels = EstimatePageContentWidthPixels(document);
                // Styling each heading would otherwise rescan the whole body for a table of contents.
                document.RunWithDeferredHeadingModified(
                    () => RenderDocument(markdown, document, options, pageContentWidthPixels, cancellationToken));
                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        private void RenderDocument(Omd.MarkdownDocument markdown, WordDocument document, MarkdownWordExportOptions options, double pageContentWidthPixels, CancellationToken cancellationToken) {
            var blocks = GetRenderableBlocks(markdown, options);

            _currentFootnotes = blocks
                .OfType<Omd.FootnoteDefinitionBlock>()
                .GroupBy(f => f.Label)
                .ToDictionary(g => g.Key, g => g.Last().Text);

            var host = new DocumentWordBlockRenderHost(document, options);
            if (markdown.DocumentHeader != null) {
                RenderSharedBlockOmd(
                    markdown.DocumentHeader,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: 0,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: Omd.MarkdownColumnAlignment.None);
            }

            foreach (var block in blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                if (block == null) {
                    continue;
                }

                RenderSharedBlockOmd(
                    block,
                    host,
                    options,
                    document,
                    currentList: null,
                    listLevel: 0,
                    quoteDepth: 0,
                    pageContentWidthPixels: pageContentWidthPixels,
                    alignment: Omd.MarkdownColumnAlignment.None);
            }

            AddSpaceAfterBodyTables(document);
        }

        private const string SpaceAfterTableTwips = "200";

        /// <summary>
        /// Word puts no space between a table and the paragraph after it, so the text sits flush against
        /// the last row. Give body paragraphs that directly follow a table at least a 10pt gap.
        /// </summary>
        private static void AddSpaceAfterBodyTables(WordDocument document) {
            var body = document._wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return;
            }

            foreach (var table in body.Elements<Table>().ToList()) {
                if (table.NextSibling() is not Paragraph paragraph) {
                    continue;
                }

                var properties = paragraph.ParagraphProperties;
                string? styleId = properties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleId.StartsWith("Heading", StringComparison.Ordinal)) {
                    continue;
                }
                if (properties?.SpacingBetweenLines?.BeforeLines != null ||
                    (int.TryParse(properties?.SpacingBetweenLines?.Before?.Value, out int before) &&
                     before >= int.Parse(SpaceAfterTableTwips, System.Globalization.CultureInfo.InvariantCulture))) {
                    continue;
                }

                properties ??= paragraph.PrependChild(new ParagraphProperties());
                properties.SpacingBetweenLines ??= new SpacingBetweenLines();
                properties.SpacingBetweenLines.Before = SpaceAfterTableTwips;
            }
        }

        private static void RemoveDuplicateNativeTocTitleHeadings(List<Omd.IMarkdownBlock> blocks) {
            for (int i = 1; i < blocks.Count; i++) {
                if (blocks[i] is not Omd.TocBlock toc ||
                    !toc.TitleHeadingAlreadyRendered ||
                    !toc.IncludeTitle ||
                    string.IsNullOrWhiteSpace(toc.Title) ||
                    toc.Scope != Omd.MarkdownTocScope.Document ||
                    blocks[i - 1] is not Omd.HeadingBlock heading) {
                    continue;
                }

                int titleLevel = toc.TitleLevel < 1 ? 1 : (toc.TitleLevel > 6 ? 6 : toc.TitleLevel);
                if (heading.Level == titleLevel &&
                    string.Equals(heading.Text.Trim(), toc.Title.Trim(), StringComparison.Ordinal)) {
                    blocks.RemoveAt(i - 1);
                    i--;
                }
            }
        }

        private IReadOnlyList<Omd.IMarkdownBlock> GetRenderableBlocks(Omd.MarkdownDocument markdown, MarkdownWordExportOptions options) {
            var (realized, catalog) = markdown.GetBlocksAndHeadingSlugs();
            var blocks = realized ?? new List<Omd.IMarkdownBlock>();
            _headingCatalog = catalog;
            options.HeadingBookmarks = BuildHeadingBookmarks(blocks, catalog);
            RemoveDuplicateNativeTocTitleHeadings(blocks);
            return blocks;
        }

        private Omd.MarkdownHeadingCatalog? _headingCatalog;

        /// <summary>
        /// Maps each heading slug (the anchor Markdown and HTML use for <c>#fragment</c> links) to a Word bookmark name.
        /// Word bookmark names allow letters, digits and underscores, must start with a letter and are limited to 40 characters.
        /// </summary>
        private static Dictionary<string, string> BuildHeadingBookmarks(IReadOnlyList<Omd.IMarkdownBlock> blocks, Omd.MarkdownHeadingCatalog catalog) {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var block in blocks) {
                if (block is not Omd.IHeadingMarkdownBlock heading) {
                    continue;
                }

                string slug = catalog.GetHeadingAnchor(heading);
                if (string.IsNullOrEmpty(slug) || map.ContainsKey(slug)) {
                    continue;
                }

                map[slug] = CreateUniqueBookmarkName(slug, used);
            }

            return map;
        }

        internal static string CreateUniqueBookmarkName(string slug, HashSet<string> used) {
            const int MaxLength = 40;
            var builder = new System.Text.StringBuilder(slug.Length);
            foreach (char c in slug) {
                builder.Append(c < 128 && char.IsLetterOrDigit(c) ? c : '_');
            }

            string name = builder.ToString().Trim('_');
            if (name.Length == 0 || !char.IsLetter(name[0])) {
                name = "h_" + name;
            }

            if (name.Length > MaxLength) {
                name = name.Substring(0, MaxLength);
            }

            string candidate = name;
            for (int suffix = 1; !used.Add(candidate); suffix++) {
                string tail = "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                candidate = (name.Length + tail.Length > MaxLength ? name.Substring(0, MaxLength - tail.Length) : name) + tail;
            }

            return candidate;
        }

        private string? ResolveHeadingBookmark(Omd.HeadingBlock heading, MarkdownWordExportOptions options) {
            if (_headingCatalog == null || options.HeadingBookmarks == null) {
                return null;
            }

            string slug = _headingCatalog.GetHeadingAnchor(heading);
            return !string.IsNullOrEmpty(slug) && options.HeadingBookmarks.TryGetValue(slug, out var name) ? name : null;
        }

        /// <summary>Returns the bookmark for an in-document <c>#fragment</c> link, or <see langword="null"/> for other links.</summary>
        internal static string? ResolveInternalLinkBookmark(string? url, MarkdownWordExportOptions options) {
            if (string.IsNullOrEmpty(url) || url![0] != '#' || url.Length < 2) {
                return null;
            }

            string fragment;
            try {
                fragment = Uri.UnescapeDataString(url.Substring(1));
            } catch (UriFormatException) {
                fragment = url.Substring(1);
            }

            if (options.HeadingBookmarks != null) {
                if (options.HeadingBookmarks.TryGetValue(fragment, out var bookmark)) {
                    return bookmark;
                }

                foreach (var pair in options.HeadingBookmarks) {
                    if (string.Equals(pair.Key, fragment, StringComparison.OrdinalIgnoreCase)) {
                        return pair.Value;
                    }
                }
            }

            // No heading owns this fragment. Keep it as an in-document link so a bookmark with that name
            // (for example one added to a template) can still satisfy it, and report the unresolved target.
            options.OnWarning?.Invoke($"Internal link '#{fragment}' does not match any heading; it was written as a Word bookmark link that may not resolve.");
            var unused = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return CreateUniqueBookmarkName(fragment, unused);
        }

    }
}
