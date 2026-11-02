using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    /// <summary>Appends offline HTML to a header and returns the created elements.</summary>
    /// <param name="header">Header to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="section">Section supplying layout geometry. Defaults to the first section defining this story; an explicit section must share or inherit it.</param>
    /// <returns>New live elements in story order, using the existing Word element projection.</returns>
    public static IReadOnlyList<WordElement> AddHtml(
        this WordHeader header,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null,
        WordSection? section = null) {
        if (header == null) throw new ArgumentNullException(nameof(header));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        WordSection target = ResolveHtmlStorySection(header, section);
        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        EnsureOfflineSynchronousImport(htmlDocument, resolved);
        return header.AddHtmlAsync(htmlDocument, resolved, target).GetAwaiter().GetResult();
    }

    /// <summary>Asynchronously appends HTML to a header and returns the created elements.</summary>
    /// <param name="header">Header to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="section">Section supplying layout geometry. Defaults to the first section defining this story; an explicit section must share or inherit it.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>New live elements in story order, using the existing Word element projection.</returns>
    public static async Task<IReadOnlyList<WordElement>> AddHtmlAsync(
        this WordHeader header,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null,
        WordSection? section = null,
        CancellationToken cancellationToken = default) {
        if (header == null) throw new ArgumentNullException(nameof(header));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        cancellationToken.ThrowIfCancellationRequested();
        WordSection target = ResolveHtmlStorySection(header, section);
        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        var converter = new HtmlToWordConverter();
        IReadOnlyList<WordElement> elements = await converter.AddHtmlToHeaderAsync(
            header.Document, target, header, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport),
            resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return elements;
    }

    /// <summary>Appends offline HTML to a footer and returns the created elements.</summary>
    /// <param name="footer">Footer to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="section">Section supplying layout geometry. Defaults to the first section defining this story; an explicit section must share or inherit it.</param>
    /// <returns>New live elements in story order, using the existing Word element projection.</returns>
    public static IReadOnlyList<WordElement> AddHtml(
        this WordFooter footer,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null,
        WordSection? section = null) {
        if (footer == null) throw new ArgumentNullException(nameof(footer));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        WordSection target = ResolveHtmlStorySection(footer, section);
        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        EnsureOfflineSynchronousImport(htmlDocument, resolved);
        return footer.AddHtmlAsync(htmlDocument, resolved, target).GetAwaiter().GetResult();
    }

    /// <summary>Asynchronously appends HTML to a footer and returns the created elements.</summary>
    /// <param name="footer">Footer to modify.</param>
    /// <param name="htmlDocument">Parsed HTML source to insert.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <param name="section">Section supplying layout geometry. Defaults to the first section defining this story; an explicit section must share or inherit it.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>New live elements in story order, using the existing Word element projection.</returns>
    public static async Task<IReadOnlyList<WordElement>> AddHtmlAsync(
        this WordFooter footer,
        HtmlConversionDocument htmlDocument,
        WordHtmlImportOptions? options = null,
        WordSection? section = null,
        CancellationToken cancellationToken = default) {
        if (footer == null) throw new ArgumentNullException(nameof(footer));
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        cancellationToken.ThrowIfCancellationRequested();
        WordSection target = ResolveHtmlStorySection(footer, section);
        WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
        resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
        var converter = new HtmlToWordConverter();
        IReadOnlyList<WordElement> elements = await converter.AddHtmlToFooterAsync(
            footer.Document, target, footer, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport),
            resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return elements;
    }

    private static WordSection ResolveHtmlStorySection(WordHeaderFooter story, WordSection? section) {
        WordDocument document = story.Document;
        if (section != null && (!ReferenceEquals(section._document, document) || !document.Sections.Contains(section))) {
            throw new ArgumentException("The layout section must belong to the receiver's document.", nameof(section));
        }
        MainDocumentPart? main = document._wordprocessingDocument?.MainDocumentPart;
        OpenXmlPart? part = story is WordHeader ? story._header?.HeaderPart : story._footer?.FooterPart;
        if (main == null || part == null || !main.Parts.Any(relationship => ReferenceEquals(relationship.OpenXmlPart, part))) {
            throw new InvalidOperationException("The header or footer is not attached to its document.");
        }
        List<WordSection> owners = WordStoryLayoutResolver.GetOwners(document, main.GetIdOfPart(part), story is WordHeader);
        if (owners.Count == 0) {
            throw new InvalidOperationException("The header or footer has no defining section in its document.");
        }
        if (section != null && !owners.Contains(section)) {
            throw new ArgumentException("The layout section must share or inherit the receiver's header or footer.", nameof(section));
        }
        return section ?? owners[0];
    }
}
