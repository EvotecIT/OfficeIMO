using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

public static partial class BookManuscriptImporter {
    /// <summary>Creates a bounded resolver for assets inside a local manuscript's physical parent directory.</summary>
    /// <remarks>The host must have permission to read these local files. Other locations and network requests are rejected.</remarks>
    public static HtmlRenderResourceResolver CreateLocalResourceResolver(string manuscriptPath, long maximumResourceBytes = 10L * 1024 * 1024) =>
        OfficeWorkflowHtmlResourceResolver.CreateResolver(manuscriptPath, maximumResourceBytes);

    /// <summary>Imports a local DOCX, Markdown or HTML file, resolving local assets only within its physical parent directory.</summary>
    public static async Task<EpubManuscriptResult> ImportFileAsync(string path, EpubManuscriptOptions? options = null,
        long maximumInputBytes = 64L * 1024 * 1024, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A manuscript path is required.", nameof(path));
        if (maximumInputBytes <= 0 || maximumInputBytes > 64L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        string fullPath = Path.GetFullPath(path);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not ".docx" and not ".md" and not ".markdown" and not ".html" and not ".htm")
            throw new NotSupportedException("Manuscript import supports DOCX, Markdown and HTML files.");
        var operation = (options ?? new EpubManuscriptOptions()).Clone();
        operation.ResourceResolver ??= OfficeWorkflowHtmlResourceResolver.CreateResolver(fullPath, operation.MaxResourceBytes);
        var htmlOptions = new HtmlConversionDocumentOptions {
            BaseUri = new Uri(fullPath), ResourceUrlPolicy = OfficeWorkflowHtmlResourceResolver.CreateResourcePolicy()
        };
        byte[] bytes = OfficeWorkflowInputReader.ReadAllBytes(fullPath, maximumInputBytes, cancellationToken);
        return await ImportBytesCoreAsync(bytes, extension, operation, htmlOptions,
            Path.GetFileNameWithoutExtension(fullPath), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Imports an application-owned DOCX, Markdown or HTML snapshot without opening files or resolving assets implicitly.</summary>
    public static async Task<EpubManuscriptResult> ImportBytesAsync(byte[] bytes, string extension, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        return await ImportBytesCoreAsync(bytes, extension, options, htmlOptions, null, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<EpubManuscriptResult> ImportBytesCoreAsync(byte[] bytes, string extension, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions, string? fallbackTitle, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(extension);
        if (bytes.LongLength > 64L * 1024 * 1024) throw new InvalidDataException("The manuscript snapshot exceeds 64 MiB.");
        extension = "." + extension.TrimStart('.').ToLowerInvariant();
        var operation = options.Clone();
        cancellationToken.ThrowIfCancellationRequested();
        if (extension == ".docx") {
            using var input = new MemoryStream(bytes, false);
            var package = OfficePackageSecurityOptions.SecureDefaults;
            package.MaxPackageBytes = 64L * 1024 * 1024;
            package.MaxPartCount = 2048;
            package.MaxPartUncompressedBytes = 32L * 1024 * 1024;
            package.MaxTotalUncompressedBytes = 128L * 1024 * 1024;
            package.MaxXmlCharactersInPart = 16L * 1024 * 1024;
            package.Macros = package.ActiveX = OfficePackageContentPolicy.Reject;
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly,
                MaxInputBytes = 64L * 1024 * 1024, PackageSecurity = package });
            if (string.IsNullOrWhiteSpace(document.BuiltinDocumentProperties.Title)) operation.Title ??= fallbackTitle;
            return await ImportWordAsync(document, operation, htmlOptions, cancellationToken).ConfigureAwait(false);
        }
        using var stream = new MemoryStream(bytes, false);
        if (extension is ".md" or ".markdown") {
            using var reader = new StreamReader(stream, new System.Text.UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: true);
            MarkdownParseResult parsed = MarkdownDoc.ParseResult(await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false));
            if (parsed.Document.FindFrontMatterEntry("title")?.Value is not string title || string.IsNullOrWhiteSpace(title))
                operation.Title ??= fallbackTitle;
            return await ImportMarkdownAsync(parsed.Document, operation, htmlOptions, cancellationToken).ConfigureAwait(false);
        }
        if (extension is not ".html" and not ".htm") throw new NotSupportedException("Manuscript snapshots support DOCX, Markdown and HTML.");
        HtmlConversionDocument html = await HtmlConversionDocument.LoadAsync(stream, htmlOptions, cancellationToken: cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(html.Document.QuerySelector("title")?.TextContent)) operation.Title ??= fallbackTitle ?? "Book";
        return await EpubManuscript.ImportHtmlAsync(html, operation, cancellationToken).ConfigureAwait(false);
    }
}
