using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Converts HTML fragments or documents into OfficeIMO.Markdown documents.
/// </summary>
public sealed partial class HtmlToMarkdownConverter {
    private sealed class ConversionContext {
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

    private static bool ShouldIgnoreElement(IElement element, ConversionContext context) {
        if (!context.Options.RemoveScriptsAndStyles) {
            return false;
        }

        string name = element.TagName;
        return name.Equals("SCRIPT", StringComparison.OrdinalIgnoreCase)
               || name.Equals("STYLE", StringComparison.OrdinalIgnoreCase)
               || name.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase)
               || name.Equals("TEMPLATE", StringComparison.OrdinalIgnoreCase);
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
