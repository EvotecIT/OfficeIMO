using System.Threading;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Providers;

/// <summary>Inert HTML parsing backed by AngleSharp, returning owned OfficeIMO document snapshots.</summary>
public sealed class AngleSharpHtmlParser : IHtmlParserProvider {
    /// <summary>Stateless default parser shared by conversion requests.</summary>
    public static AngleSharpHtmlParser Instance { get; } = new AngleSharpHtmlParser();
    /// <inheritdoc />
    public string Id => "AngleSharp/" + typeof(global::AngleSharp.Html.Parser.HtmlParser).Assembly.GetName().Version;

    /// <inheritdoc />
    public HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        HtmlParseOptions resolved = (options ?? throw new ArgumentNullException(nameof(options))).Clone();
        resolved.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        if (resolved.MaxInputCharacters.HasValue && source.Length > resolved.MaxInputCharacters.Value) throw new HtmlParseLimitException(nameof(options.MaxInputCharacters), source.Length, resolved.MaxInputCharacters.Value);
        var native = HtmlDocumentParser.ParseDocument(source, cancellationToken);
        return NativeDomBridge.Import(native, resolved, cancellationToken).Freeze();
    }
}
