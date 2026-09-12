using AngleSharp.Dom;
using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Context supplied to HTML element block converters.
/// </summary>
public sealed class HtmlElementBlockConversionContext {
    private readonly Func<IEnumerable<INode>, IReadOnlyList<IMarkdownBlock>> _convertNodesToBlocks;
    private readonly Func<IEnumerable<INode>, InlineSequence> _convertNodesToInlineSequence;
    private readonly Func<string?, string> _normalizeBlockText;
    private readonly IElement _nativeElement;
    private HtmlElement? _element;

    internal HtmlElementBlockConversionContext(
        IElement element,
        HtmlToMarkdownOptions options,
        Func<IEnumerable<INode>, IReadOnlyList<IMarkdownBlock>> convertNodesToBlocks,
        Func<IEnumerable<INode>, InlineSequence> convertNodesToInlineSequence,
        Func<string?, string> normalizeBlockText) {
        _nativeElement = element ?? throw new ArgumentNullException(nameof(element));
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _convertNodesToBlocks = convertNodesToBlocks ?? throw new ArgumentNullException(nameof(convertNodesToBlocks));
        _convertNodesToInlineSequence = convertNodesToInlineSequence ?? throw new ArgumentNullException(nameof(convertNodesToInlineSequence));
        _normalizeBlockText = normalizeBlockText ?? throw new ArgumentNullException(nameof(normalizeBlockText));
    }

    /// <summary>Current HTML element being converted.</summary>
    public HtmlElement Element => _element ??= NativeDomBridge.Wrap(_nativeElement);

    /// <summary>Effective HTML-to-markdown options.</summary>
    public HtmlToMarkdownOptions Options { get; }

    /// <summary>Converts the supplied nodes using the base block converter, skipping null entries.</summary>
    public IReadOnlyList<IMarkdownBlock> ConvertNodesToBlocks(IEnumerable<HtmlNode?> nodes) => _convertNodesToBlocks(ToNativeNodes(nodes));

    /// <summary>Converts the current element's child nodes using the base block converter.</summary>
    public IReadOnlyList<IMarkdownBlock> ConvertChildNodesToBlocks() => _convertNodesToBlocks(_nativeElement.ChildNodes);

    /// <summary>Converts the supplied nodes using the base inline converter, skipping null entries.</summary>
    public InlineSequence ConvertNodesToInlineSequence(IEnumerable<HtmlNode?> nodes) => _convertNodesToInlineSequence(ToNativeNodes(nodes));

    /// <summary>Converts the current element's child nodes using the base inline converter.</summary>
    public InlineSequence ConvertChildNodesToInlineSequence() => _convertNodesToInlineSequence(_nativeElement.ChildNodes);

    private IEnumerable<INode> ToNativeNodes(IEnumerable<HtmlNode?> nodes) =>
        (nodes ?? throw new ArgumentNullException(nameof(nodes))).Where(node => node != null)
            .Select(node => NativeDomBridge.GetCallbackNative(node!, Element.Document));

    /// <summary>Normalizes HTML text content using the base block text rules.</summary>
    public string NormalizeBlockText(string? value) => _normalizeBlockText(value);
}

/// <summary>
/// Ordered host/plugin seam for reinterpreting custom HTML elements during HTML-to-markdown conversion.
/// </summary>
public sealed class HtmlElementBlockConverter {
    private readonly Func<HtmlElementBlockConversionContext, IReadOnlyList<IMarkdownBlock>?> _convert;

    /// <summary>Create a new HTML element block converter.</summary>
    public HtmlElementBlockConverter(
        string id,
        string name,
        Func<HtmlElementBlockConversionContext, IReadOnlyList<IMarkdownBlock>?> convert) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Converter id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Converter name is required.", nameof(name));
        }

        Id = id.Trim();
        Name = name.Trim();
        _convert = convert ?? throw new ArgumentNullException(nameof(convert));
    }

    /// <summary>Stable converter identifier used for deduplication and diagnostics.</summary>
    public string Id { get; }

    /// <summary>Friendly converter name used for diagnostics and documentation.</summary>
    public string Name { get; }

    /// <summary>Attempts to create blocks from the supplied HTML element conversion context.</summary>
    public IReadOnlyList<IMarkdownBlock>? TryConvert(HtmlElementBlockConversionContext context) {
        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        return _convert(context);
    }
}
