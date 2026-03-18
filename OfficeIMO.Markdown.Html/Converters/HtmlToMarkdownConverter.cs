using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Converts HTML fragments or documents into OfficeIMO.Markdown documents.
/// </summary>
public sealed partial class HtmlToMarkdownConverter {
    internal sealed class ConversionContext {
        public ConversionContext(HtmlToMarkdownOptions options) {
            Options = options ?? throw new ArgumentNullException(nameof(options));
        }

        public HtmlToMarkdownOptions Options { get; }
    }

    /// <summary>
    /// Converts HTML into Markdown text.
    /// </summary>
    /// <param name="html">HTML fragment or document to convert.</param>
    /// <param name="options">
    /// Optional conversion options. When omitted, <see cref="HtmlToMarkdownOptions"/> defaults are used.
    /// </param>
    /// <returns>Markdown text produced from the supplied HTML.</returns>
    public string Convert(string html, HtmlToMarkdownOptions? options = null) {
        var effectiveOptions = options?.Clone() ?? new HtmlToMarkdownOptions();
        return ConvertToDocument(html, effectiveOptions).ToMarkdown(effectiveOptions.MarkdownWriteOptions);
    }

    /// <summary>
    /// Converts HTML into a Markdown document model.
    /// </summary>
    /// <param name="html">HTML fragment or document to convert.</param>
    /// <param name="options">
    /// Optional conversion options. When omitted, <see cref="HtmlToMarkdownOptions"/> defaults are used.
    /// </param>
    /// <returns>
    /// A <see cref="MarkdownDoc"/> that preserves the converted block structure before rendering to Markdown text.
    /// </returns>
    public MarkdownDoc ConvertToDocument(string html, HtmlToMarkdownOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        var effectiveOptions = options?.Clone() ?? new HtmlToMarkdownOptions();
        ValidateInputLength(html, effectiveOptions.MaxInputCharacters, nameof(html));

        var parser = new HtmlParser();
        var document = parser.ParseDocument(html);
        effectiveOptions.BaseUri = ResolveEffectiveBaseUri(document, effectiveOptions.BaseUri);
        var context = new ConversionContext(effectiveOptions);

        INode root = effectiveOptions.UseBodyContentsOnly && document.Body != null
            ? document.Body
            : (INode?)document.DocumentElement ?? document;

        var markdown = MarkdownDoc.Create();
        foreach (var block in ConvertNodesToBlocks(root.ChildNodes, context)) {
            markdown.Add(block);
        }

        return MarkdownDocumentTransformPipeline.Apply(
            markdown,
            effectiveOptions.DocumentTransforms,
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.HtmlToMarkdown, effectiveOptions));
    }

    private static Uri? ResolveEffectiveBaseUri(IHtmlDocument document, Uri? fallbackBaseUri) {
        if (document == null) {
            return fallbackBaseUri;
        }

        var baseElement = document.QuerySelector("base[href]");
        string? rawBaseHref = baseElement?.GetAttribute("href");
        if (rawBaseHref == null) {
            return fallbackBaseUri;
        }

        string baseHref = rawBaseHref.Trim();
        if (baseHref.Length == 0) {
            return fallbackBaseUri;
        }

        if (fallbackBaseUri != null && Uri.TryCreate(fallbackBaseUri, baseHref, out var resolvedFromFallback)) {
            return resolvedFromFallback;
        }

        return Uri.TryCreate(baseHref, UriKind.Absolute, out var absoluteBaseUri)
            ? absoluteBaseUri
            : fallbackBaseUri;
    }

    private static bool ShouldIgnoreElement(IElement element, ConversionContext context) {
        if (!context.Options.RemoveScriptsAndStyles) {
            return ShouldSuppressListingCardMetadataElement(element, context);
        }

        string name = element.TagName;
        return name.Equals("SCRIPT", StringComparison.OrdinalIgnoreCase)
               || name.Equals("STYLE", StringComparison.OrdinalIgnoreCase)
               || name.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase)
               || name.Equals("TEMPLATE", StringComparison.OrdinalIgnoreCase)
               || ShouldSuppressListingCardMetadataElement(element, context);
    }

    private static bool ShouldSuppressListingCardMetadataElement(IElement element, ConversionContext context) {
        if (element == null
            || context?.Options == null
            || context.Options.ListingCardMetadataMode != HtmlListingCardMetadataMode.SuppressInRepeatedCards) {
            return false;
        }

        return LooksLikeListingCardMetadataElement(element)
               && TryFindRepeatedListingCardRoot(element, out _);
    }

    private static bool LooksLikeListingCardMetadataElement(IElement element) {
        if (element == null) {
            return false;
        }

        string signalText = string.Join(" ",
            element.TagName,
            element.Id ?? string.Empty,
            element.ClassName ?? string.Empty,
            element.GetAttribute("role") ?? string.Empty,
            element.GetAttribute("data-testid") ?? string.Empty,
            element.GetAttribute("itemprop") ?? string.Empty).ToLowerInvariant();

        if (signalText.Contains("post-date", StringComparison.Ordinal)
            || signalText.Contains("post-time", StringComparison.Ordinal)
            || signalText.Contains("entry-meta", StringComparison.Ordinal)
            || signalText.Contains("post-meta", StringComparison.Ordinal)
            || signalText.Contains("meta-author", StringComparison.Ordinal)
            || signalText.Contains("post-author", StringComparison.Ordinal)
            || signalText.Contains("post-meta-author", StringComparison.Ordinal)
            || signalText.Contains("post-meta-categories", StringComparison.Ordinal)
            || signalText.Contains("post-categories", StringComparison.Ordinal)
            || signalText.Contains("post-read-more", StringComparison.Ordinal)
            || signalText.Contains("read-more", StringComparison.Ordinal)
            || signalText.Contains("more-link", StringComparison.Ordinal)
            || signalText.Contains("post-links", StringComparison.Ordinal)
            || signalText.Contains("post-misc", StringComparison.Ordinal)
            || signalText.Contains("comment-link", StringComparison.Ordinal)
            || signalText.Contains("post-comments", StringComparison.Ordinal)) {
            return true;
        }

        if (element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)) {
            string inlineText = NormalizeBlockText(element.TextContent);
            if (inlineText.Equals("Read More", StringComparison.OrdinalIgnoreCase)
                || inlineText.Equals("Continue Reading", StringComparison.OrdinalIgnoreCase)
                || inlineText.Equals("Continue reading", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryFindRepeatedListingCardRoot(IElement element, out IElement? cardRoot) {
        cardRoot = null;
        for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
            if (!LooksLikeListingCardRoot(current)) {
                continue;
            }

            if (HasRepeatedListingCardSiblings(current)) {
                cardRoot = current;
                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeListingCardRoot(IElement element) {
        if (element == null) {
            return false;
        }

        string tag = element.TagName;
        if (!tag.Equals("ARTICLE", StringComparison.OrdinalIgnoreCase)
            && !tag.Equals("DIV", StringComparison.OrdinalIgnoreCase)
            && !tag.Equals("SECTION", StringComparison.OrdinalIgnoreCase)
            && !tag.Equals("LI", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        bool hasTitleLink = element.QuerySelector("h1 a[href], h2 a[href], h3 a[href], h4 a[href], .entry-title a[href], .post-title a[href], a[rel='bookmark']") != null;
        bool hasMediaOrSummary = element.QuerySelector("img, picture, .summary, .excerpt, p") != null;
        if (!hasTitleLink || !hasMediaOrSummary) {
            return false;
        }

        if (tag.Equals("ARTICLE", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string signalText = string.Join(" ",
            element.Id ?? string.Empty,
            element.ClassName ?? string.Empty).ToLowerInvariant();

        return signalText.Contains("post", StringComparison.Ordinal)
               || signalText.Contains("entry", StringComparison.Ordinal)
               || signalText.Contains("card", StringComparison.Ordinal)
               || signalText.Contains("teaser", StringComparison.Ordinal)
               || signalText.Contains("listing", StringComparison.Ordinal)
               || signalText.Contains("timeline", StringComparison.Ordinal)
               || signalText.Contains("blog", StringComparison.Ordinal)
               || signalText.Contains("feed", StringComparison.Ordinal)
               || signalText.Contains("item", StringComparison.Ordinal);
    }

    private static bool HasRepeatedListingCardSiblings(IElement candidate) {
        IElement? parent = candidate.ParentElement;
        if (parent == null) {
            return false;
        }

        int siblingCardCount = 0;
        foreach (IElement sibling in parent.Children) {
            if (!LooksLikeListingCardRoot(sibling)) {
                continue;
            }

            siblingCardCount++;
            if (siblingCardCount >= 2) {
                return true;
            }
        }

        return false;
    }

    private static string ResolveUrl(string? rawUrl, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(rawUrl)) {
            return string.Empty;
        }

        string candidate = rawUrl!.Trim();
        if (context.Options.BaseUri == null) {
            return candidate;
        }

        if (Uri.TryCreate(context.Options.BaseUri, candidate, out var resolved)) {
            return resolved.AbsoluteUri;
        }

        return candidate;
    }

    private static string NormalizeBlockText(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        return CollapseWhitespace(value!).Trim();
    }

    private static void ValidateInputLength(string input, int? maxInputCharacters, string paramName) {
        if (!maxInputCharacters.HasValue) {
            return;
        }

        if (maxInputCharacters.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxInputCharacters), maxInputCharacters.Value, "MaxInputCharacters must be greater than zero.");
        }

        if (input.Length > maxInputCharacters.Value) {
            throw new ArgumentOutOfRangeException(paramName, input.Length, $"Input exceeds MaxInputCharacters ({maxInputCharacters.Value}).");
        }
    }

    private static string RenderBlocksToMarkdown(IEnumerable<IMarkdownBlock> blocks) {
        var renderedBlocks = new List<string>();
        foreach (var block in blocks) {
            string rendered = block.RenderMarkdown().Trim();
            if (rendered.Length > 0) {
                renderedBlocks.Add(rendered);
            }
        }

        return string.Join("\n\n", renderedBlocks);
    }

    private static string CollapseWhitespace(string value) {
        var sb = new StringBuilder(value.Length);
        bool previousWhitespace = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            bool isWhitespace = char.IsWhiteSpace(ch);
            if (isWhitespace) {
                if (!previousWhitespace) {
                    sb.Append(' ');
                }
                previousWhitespace = true;
            } else {
                sb.Append(ch);
                previousWhitespace = false;
            }
        }
        return sb.ToString();
    }

    private static string EscapeInlineText(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        var sb = new StringBuilder(value!.Length + 8);
        bool previousWhitespace = false;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsWhiteSpace(ch)) {
                if (!previousWhitespace) {
                    sb.Append(' ');
                }
                previousWhitespace = true;
                continue;
            }

            previousWhitespace = false;
            switch (ch) {
                case '\\':
                case '`':
                case '*':
                case '_':
                case '[':
                case ']':
                case '|':
                    sb.Append('\\').Append(ch);
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }

        return sb.ToString();
    }

    private static string EscapeLinkTarget(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var target = value!.Trim();
        return target
            .Replace(" ", "%20")
            .Replace("(", "%28")
            .Replace(")", "%29");
    }

    private static string WrapCode(string? value) {
        string code = value ?? string.Empty;
        int longestRun = 0;
        int currentRun = 0;
        for (int i = 0; i < code.Length; i++) {
            if (code[i] == '`') {
                currentRun++;
                if (currentRun > longestRun) {
                    longestRun = currentRun;
                }
            } else {
                currentRun = 0;
            }
        }

        string fence = new string('`', longestRun + 1);
        return fence + code + fence;
    }

    private static InlineSequence ParseInlines(string? markdown) {
        string normalized = NormalizeBlockText(markdown);
        if (normalized.Length == 0) {
            return new InlineSequence();
        }

        return MarkdownReader.ParseInlineText(normalized);
    }
}
