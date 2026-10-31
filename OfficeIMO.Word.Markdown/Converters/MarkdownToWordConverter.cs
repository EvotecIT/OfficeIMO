using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
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

        public WordDocument Convert(Omd.MarkdownDocument markdown, WordMarkdownImportOptions options, CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new WordMarkdownImportOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);
            ApplyThemeHeadingStyles(document, OfficeTheme.Default);
            var pageContentWidthPixels = EstimatePageContentWidthPixels(document);

            RenderDocument(markdown, document, options, pageContentWidthPixels, cancellationToken);
            return document;
        }

        private void RenderDocument(Omd.MarkdownDocument markdown, WordDocument document, WordMarkdownImportOptions options, double pageContentWidthPixels, CancellationToken cancellationToken) {
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
                if (styleId != null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) {
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

        /// <summary>
        /// Gives a newly created document a clear heading hierarchy from the theme: heading font, type-scale
        /// sizes, bold, the text colour and keep-with-next. Applied to new documents only, never to templates.
        /// </summary>
        private static void ApplyThemeHeadingStyles(WordDocument document, OfficeTheme theme) {
            document.EnsureStyleDefinitionsInitialized();
            var styles = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return;
            }

            OfficeLength[] sizes = {
                theme.TypeScale.Heading1, theme.TypeScale.Heading2, theme.TypeScale.Heading3,
                theme.TypeScale.Heading4, theme.TypeScale.Heading5, theme.TypeScale.Heading6
            };
            string font = theme.Fonts.Heading.Family;
            string color = theme.Palette.Text.ToRgbHex();
            for (int level = 1; level <= sizes.Length; level++) {
                string styleId = "Heading" + level.ToString(System.Globalization.CultureInfo.InvariantCulture);
                Style? style = styles.Elements<Style>().FirstOrDefault(item => string.Equals(item.StyleId?.Value, styleId, StringComparison.Ordinal));
                if (style == null) {
                    continue;
                }

                StyleParagraphProperties paragraphProperties = style.StyleParagraphProperties ??= new StyleParagraphProperties();
                if (paragraphProperties.GetFirstChild<KeepNext>() == null) {
                    paragraphProperties.PrependChild(new KeepNext());
                }

                StyleRunProperties runProperties = style.StyleRunProperties ??= new StyleRunProperties();
                string halfPoints = Math.Round(sizes[level - 1].ToPoints() * 2D, MidpointRounding.AwayFromZero)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                SetStyleRunChild(runProperties, new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font });
                SetStyleRunChild(runProperties, new Bold());
                SetStyleRunChild(runProperties, new BoldComplexScript());
                SetStyleRunChild(runProperties, new Color { Val = color });
                SetStyleRunChild(runProperties, new FontSize { Val = halfPoints });
                SetStyleRunChild(runProperties, new FontSizeComplexScript { Val = halfPoints });
            }

            // Normal: body font, body size and the text colour, with the theme's paragraph spacing.
            Style? normal = styles.Elements<Style>().FirstOrDefault(item => string.Equals(item.StyleId?.Value, "Normal", StringComparison.Ordinal));
            if (normal != null) {
                string bodyFont = theme.Fonts.Body.Family;
                string bodySize = Math.Round(theme.TypeScale.Body.ToPoints() * 2D, MidpointRounding.AwayFromZero)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                StyleRunProperties normalRun = normal.StyleRunProperties ??= new StyleRunProperties();
                SetStyleRunChild(normalRun, new RunFonts { Ascii = bodyFont, HighAnsi = bodyFont, ComplexScript = bodyFont, EastAsia = bodyFont });
                SetStyleRunChild(normalRun, new Color { Val = color });
                SetStyleRunChild(normalRun, new FontSize { Val = bodySize });
                SetStyleRunChild(normalRun, new FontSizeComplexScript { Val = bodySize });
                StyleParagraphProperties normalParagraph = normal.StyleParagraphProperties ??= new StyleParagraphProperties();
                SpacingBetweenLines spacing = normalParagraph.GetFirstChild<SpacingBetweenLines>() ?? normalParagraph.AppendChild(new SpacingBetweenLines());
                spacing.After = Math.Round(theme.Spacing.ParagraphAfter.ToPoints() * 20D, MidpointRounding.AwayFromZero)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        // CT_RPr child sequence; children are re-ordered after edits so the style stays schema-valid.
        private static readonly Type[] RunPropertyOrder = {
            typeof(RunStyle), typeof(RunFonts), typeof(Bold), typeof(BoldComplexScript), typeof(Italic), typeof(ItalicComplexScript),
            typeof(Caps), typeof(SmallCaps), typeof(Strike), typeof(DoubleStrike), typeof(Outline), typeof(Shadow), typeof(Emboss),
            typeof(Imprint), typeof(NoProof), typeof(SnapToGrid), typeof(Vanish), typeof(WebHidden), typeof(Color), typeof(Spacing),
            typeof(CharacterScale), typeof(Kern), typeof(Position), typeof(FontSize), typeof(FontSizeComplexScript), typeof(Highlight),
            typeof(Underline), typeof(TextEffect), typeof(Border), typeof(Shading), typeof(FitText), typeof(VerticalTextAlignment),
            typeof(RightToLeftText), typeof(ComplexScript), typeof(Emphasis), typeof(Languages), typeof(EastAsianLayout), typeof(SpecVanish)
        };

        private static void SetStyleRunChild<T>(StyleRunProperties runProperties, T value) where T : DocumentFormat.OpenXml.OpenXmlElement {
            foreach (T existing in runProperties.Elements<T>().ToList()) {
                existing.Remove();
            }

            runProperties.AppendChild(value);
            var children = runProperties.ChildElements.ToList();
            int Rank(DocumentFormat.OpenXml.OpenXmlElement element) {
                int index = Array.IndexOf(RunPropertyOrder, element.GetType());
                return index < 0 ? RunPropertyOrder.Length : index;
            }

            var ordered = children.Select((element, position) => (element, position))
                .OrderBy(item => Rank(item.element)).ThenBy(item => item.position).Select(item => item.element).ToList();
            runProperties.RemoveAllChildren();
            foreach (var element in ordered) {
                runProperties.AppendChild(element);
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

        private IReadOnlyList<Omd.IMarkdownBlock> GetRenderableBlocks(Omd.MarkdownDocument markdown, WordMarkdownImportOptions options) {
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

        private string? ResolveHeadingBookmark(Omd.HeadingBlock heading, WordMarkdownImportOptions options) {
            if (_headingCatalog == null || options.HeadingBookmarks == null) {
                return null;
            }

            string slug = _headingCatalog.GetHeadingAnchor(heading);
            return !string.IsNullOrEmpty(slug) && options.HeadingBookmarks.TryGetValue(slug, out var name) ? name : null;
        }

        /// <summary>Returns the bookmark for an in-document <c>#fragment</c> link, or <see langword="null"/> for other links.</summary>
        internal static string? ResolveInternalLinkBookmark(string? url, WordMarkdownImportOptions options) {
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
