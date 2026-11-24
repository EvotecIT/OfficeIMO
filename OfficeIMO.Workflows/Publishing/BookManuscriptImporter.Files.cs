using OfficeIMO.Epub;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

public static partial class BookManuscriptImporter {
    /// <summary>Creates a bounded resolver for assets inside a local manuscript's physical parent directory.</summary>
    /// <remarks>The host must have permission to read these local files. Other locations and network requests are rejected.</remarks>
    public static HtmlRenderResourceResolver CreateLocalResourceResolver(string manuscriptPath, long maximumResourceBytes = 10L * 1024 * 1024) =>
        OfficeHtmlResourceResolver.CreateResolver(manuscriptPath, maximumResourceBytes);

    /// <summary>Imports a local DOCX, Markdown or HTML file, resolving local assets only within its physical parent directory.</summary>
    public static async Task<EpubManuscriptResult> ImportFileAsync(string path, EpubManuscriptOptions? options = null,
        long maximumInputBytes = 64L * 1024 * 1024, CancellationToken cancellationToken = default) {
        if ((path == null || string.IsNullOrWhiteSpace(path))) throw new ArgumentException("A manuscript path is required.", nameof(path));
        if (maximumInputBytes <= 0 || maximumInputBytes > 64L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        string fullPath = Path.GetFullPath(path);
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not ".docx" and not ".md" and not ".markdown" and not ".html" and not ".htm")
            throw new NotSupportedException("Manuscript import supports DOCX, Markdown and HTML files.");
        var operation = (options ?? new EpubManuscriptOptions()).Clone();
        operation.ResourceResolver ??= OfficeHtmlResourceResolver.CreateResolver(fullPath, operation.MaxResourceBytes);
        var htmlOptions = new HtmlConversionDocumentOptions {
            BaseUri = new Uri(fullPath), ResourceUrlPolicy = OfficeHtmlResourceResolver.CreateResourcePolicy()
        };
        byte[] bytes = OfficeWorkflowInputReader.ReadAllBytes(fullPath, maximumInputBytes, cancellationToken);
        return await ImportBytesCoreAsync(bytes, extension, operation, htmlOptions,
            Path.GetFileNameWithoutExtension(fullPath), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Imports an application-owned DOCX, Markdown or HTML snapshot without opening files or resolving assets implicitly.</summary>
    public static Task<EpubManuscriptResult> ImportBytesAsync(byte[] bytes, string extension, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        return ImportBytesCoreAsync(bytes, extension, options, htmlOptions, null, cancellationToken);
    }

    private static async Task<EpubManuscriptResult> ImportBytesCoreAsync(byte[] bytes, string extension, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions, string? fallbackTitle, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.LongLength > MaximumManuscriptBytes) throw new InvalidDataException("The manuscript snapshot exceeds 64 MiB.");
        using var stream = new MemoryStream(bytes, false);
        return await ImportStreamAsync(stream, extension, options, htmlOptions, fallbackTitle, forWrite: false, cancellationToken).ConfigureAwait(false);
    }

    private const long MaximumManuscriptBytes = 64L * 1024 * 1024;

    internal static async Task<EpubManuscriptResult> ImportStreamAsync(Stream stream, string extension, EpubManuscriptOptions options,
        HtmlConversionDocumentOptions? htmlOptions, string? fallbackTitle, bool forWrite, CancellationToken cancellationToken,
        long maximumXmlCharactersInPart = 16L * 1024 * 1024) {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(extension);
        if (!stream.CanRead || !stream.CanSeek)
            throw new ArgumentException("The manuscript reader requires a captured seekable input.", nameof(stream));
        if (stream.Length - stream.Position > MaximumManuscriptBytes) throw new InvalidDataException("The manuscript snapshot exceeds 64 MiB.");
        extension = "." + extension.TrimStart('.').ToLowerInvariant();
        var operation = options.Clone();
        cancellationToken.ThrowIfCancellationRequested();
        if (extension == ".docx") {
            var package = OfficePackageSecurityOptions.SecureDefaults;
            package.MaxPackageBytes = 64L * 1024 * 1024;
            package.MaxPartCount = 2048;
            package.MaxPartUncompressedBytes = 32L * 1024 * 1024;
            package.MaxTotalUncompressedBytes = 128L * 1024 * 1024;
            package.MaxXmlCharactersInPart = Math.Min(maximumXmlCharactersInPart, 16L * 1024 * 1024);
            package.Macros = package.ActiveX = OfficePackageContentPolicy.Reject;
            using WordDocument document = WordDocument.Load(stream, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly,
                MaxInputBytes = MaximumManuscriptBytes, PackageSecurity = package,
                OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = package.MaxXmlCharactersInPart } });
            if (!forWrite && string.IsNullOrWhiteSpace(document.Properties.Title)) operation.Title ??= fallbackTitle;
            return await ImportWordCoreAsync(document, operation, htmlOptions, forWrite ? fallbackTitle ?? "Book" : null,
                cancellationToken).ConfigureAwait(false);
        }
        if (extension is ".md" or ".markdown") {
            using var reader = new StreamReader(stream, new System.Text.UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: true);
            // The manuscript is rendered to HTML once, so the syntax tree is not captured.
            using MarkdownDocument parsed = MarkdownDocument.Parse(await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false),
                new MarkdownLoadOptions { CaptureSyntaxTree = false });
            if (!forWrite && string.IsNullOrWhiteSpace(MarkdownFrontMatterText.Get(parsed, "title"))) operation.Title ??= fallbackTitle;
            return await ImportMarkdownCoreAsync(parsed, operation, htmlOptions, forWrite ? fallbackTitle ?? "Book" : null,
                cancellationToken).ConfigureAwait(false);
        }
        if (extension is not ".html" and not ".htm") throw new NotSupportedException("Manuscript snapshots support DOCX, Markdown and HTML.");
        HtmlConversionDocument html = await HtmlConversionDocument.LoadAsync(stream, htmlOptions, cancellationToken: cancellationToken).ConfigureAwait(false);
        if (!forWrite && string.IsNullOrWhiteSpace(html.Document.QuerySelector("title")?.TextContent)) operation.Title ??= fallbackTitle ?? "Book";
        return await ImportPreparedAsync(html, operation, forWrite ? fallbackTitle ?? "Book" : null, useHtmlTitle: true,
            cancellationToken).ConfigureAwait(false);
    }
}
