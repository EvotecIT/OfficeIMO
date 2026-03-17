using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Context passed to custom HTML inline-element converters.
/// </summary>
public sealed class HtmlInlineElementConversionContext {
    private readonly HtmlToMarkdownConverter.ConversionContext? _conversionContext;

    internal HtmlInlineElementConversionContext(IElement element, HtmlToMarkdownOptions options, HtmlToMarkdownConverter.ConversionContext? conversionContext) {
        Element = element ?? throw new ArgumentNullException(nameof(element));
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _conversionContext = conversionContext;
    }

    /// <summary>
    /// HTML element being converted.
    /// </summary>
    public IElement Element { get; }

    /// <summary>
    /// Active HTML-to-markdown options.
    /// </summary>
    public HtmlToMarkdownOptions Options { get; }

    /// <summary>
    /// Converts the supplied HTML nodes into an inline markdown sequence using the current conversion profile.
    /// </summary>
    public InlineSequence ConvertNodesToInlineSequence(IEnumerable<INode> nodes) {
        return HtmlToMarkdownConverter.ConvertInlineNodesToInlineSequence(nodes ?? Array.Empty<INode>(), _conversionContext);
    }

    /// <summary>
    /// Converts the current element's children into an inline markdown sequence using the current conversion profile.
    /// </summary>
    public InlineSequence ConvertChildNodesToInlineSequence() {
        return ConvertNodesToInlineSequence(Element.ChildNodes);
    }

    /// <summary>
    /// Normalizes plain inline text using HTML-style collapsed whitespace rules.
    /// </summary>
    public string NormalizeInlineText(string? value) {
        return HtmlToMarkdownConverter.CollapseHtmlInlineWhitespace(value ?? string.Empty).Trim();
    }
}

/// <summary>
/// Custom inline HTML element decoder used during HTML-to-markdown conversion.
/// </summary>
public sealed class HtmlInlineElementConverter {
    private readonly Func<HtmlInlineElementConversionContext, IReadOnlyList<IMarkdownInline>?> _converter;

    /// <summary>
    /// Creates a new custom inline HTML element converter.
    /// </summary>
    public HtmlInlineElementConverter(
        string id,
        string name,
        Func<HtmlInlineElementConversionContext, IReadOnlyList<IMarkdownInline>?> converter) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Converter id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Converter name is required.", nameof(name));
        }

        _converter = converter ?? throw new ArgumentNullException(nameof(converter));
        Id = id.Trim();
        Name = name.Trim();
    }

    /// <summary>
    /// Stable converter identifier used for idempotence and diagnostics.
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// Friendly converter name used for diagnostics and documentation.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Attempts to convert a matching HTML inline element into markdown inline nodes.
    /// </summary>
    public IReadOnlyList<IMarkdownInline>? Convert(HtmlInlineElementConversionContext context) {
        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        return _converter(context);
    }
}
