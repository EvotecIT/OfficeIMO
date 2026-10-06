using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Provider-neutral front door for inert HTML document and contextual-fragment parsing.
/// Returned trees are owned by OfficeIMO and do not expose provider nodes.
/// </summary>
public sealed class HtmlDocumentEngine {
    /// <summary>The default engine shipped by OfficeIMO.Html.</summary>
    public static HtmlDocumentEngine Default { get; } =
        new HtmlDocumentEngine(Providers.AngleSharpHtmlParser.Instance);

    /// <summary>Creates an engine over a caller-selected parser provider.</summary>
    public HtmlDocumentEngine(IHtmlParserProvider parserProvider) {
        ParserProvider = parserProvider ?? throw new ArgumentNullException(nameof(parserProvider));
    }

    /// <summary>The selected inert parser provider.</summary>
    public IHtmlParserProvider ParserProvider { get; }

    /// <summary>Parses a complete HTML document into an immutable owned snapshot.</summary>
    public HtmlDocument ParseDocument(
        string source,
        HtmlParseOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        HtmlParseOptions resolved = options?.Clone() ?? new HtmlParseOptions();
        resolved.Validate();
        return ParserProvider.ParseDocument(source, resolved, cancellationToken)
            ?? throw new InvalidOperationException("The HTML parser provider returned no document.");
    }

    /// <summary>
    /// Parses source using an owned element and its ancestors as the HTML fragment context.
    /// The result belongs to an independent immutable document. Import it into a mutable target
    /// document with <see cref="HtmlDocument.ImportNode"/> before insertion.
    /// </summary>
    public HtmlDocumentFragment ParseFragment(
        string source,
        HtmlElement contextElement,
        HtmlParseOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (contextElement == null) throw new ArgumentNullException(nameof(contextElement));
        HtmlParseOptions resolved = options?.Clone() ?? new HtmlParseOptions();
        resolved.Validate();
        return ParserProvider.ParseFragment(source, contextElement, resolved, cancellationToken)
            ?? throw new InvalidOperationException("The HTML parser provider returned no fragment.");
    }
}
