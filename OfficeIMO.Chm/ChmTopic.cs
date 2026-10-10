namespace OfficeIMO.Chm;

/// <summary>An HTML topic in deterministic contents order, followed by unlisted topics.</summary>
public sealed class ChmTopic {
    private readonly ChmDocument _book;
    internal ChmTopic(ChmDocument book, ChmEntry entry, string title) { _book = book; Entry = entry; Title = title; }
    /// <summary>The original HTML entry. Raw bytes remain available without text normalization.</summary>
    public ChmEntry Entry { get; }
    /// <summary>Canonical archive path.</summary>
    public string Path => Entry.Path;
    /// <summary>Compiled or sitemap title, falling back to the archive path.</summary>
    public string Title { get; }
    /// <summary>Decodes source HTML using BOM, charset, and help-book locale information.</summary>
    public string ReadHtml(CancellationToken cancellationToken = default) => _book.ReadHtml(Entry, cancellationToken);
    /// <summary>Creates the canonical inert HTML conversion document for this topic.</summary>
    public HtmlConversionDocument ToHtmlDocument(CancellationToken cancellationToken = default) =>
        HtmlConversionDocument.Parse(ReadHtml(cancellationToken), _book.CreateHtmlOptions(Path), cancellationToken);
}
