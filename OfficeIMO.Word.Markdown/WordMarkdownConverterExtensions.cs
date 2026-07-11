using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Extension methods enabling Markdown conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static class WordMarkdownConverterExtensions {
        /// <summary>
        /// Synchronously saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsMarkdown(this WordDocument document, string path, WordToMarkdownOptions? options = null) {
            document.SaveAsMarkdownAsync(path, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Synchronously saves the document as Markdown to the provided stream.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsMarkdown(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null) {
            document.SaveAsMarkdownAsync(stream, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A task representing the asynchronous save operation.</returns>
        public static async Task SaveAsMarkdownAsync(this WordDocument document, string path, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            var effectiveOptions = CreateFileSaveOptions(path, options);
            var markdown = await document.ToMarkdownAsync(effectiveOptions, cancellationToken).ConfigureAwait(false);
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_0_OR_GREATER
            await File.WriteAllTextAsync(path, markdown, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
#else
            using var writer = new StreamWriter(path, false, Encoding.UTF8);
            await writer.WriteAsync(markdown).ConfigureAwait(false);
#endif
        }

        /// <summary>
        /// Asynchronously saves the document as Markdown to the provided stream.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A task representing the asynchronous save operation.</returns>
        public static async Task SaveAsMarkdownAsync(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            options ??= new WordToMarkdownOptions();
            var markdown = await document.ToMarkdownAsync(options, cancellationToken).ConfigureAwait(false);
            var bytes = Encoding.UTF8.GetBytes(markdown);
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_0_OR_GREATER
            await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
#else
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
        }

        private static WordToMarkdownOptions CreateFileSaveOptions(string markdownPath, WordToMarkdownOptions? options) {
            var effectiveOptions = CopyOptions(options ?? new WordToMarkdownOptions());
            if (effectiveOptions.ImageExportMode == ImageExportMode.File && string.IsNullOrEmpty(effectiveOptions.ImageDirectory)) {
                effectiveOptions.ImageDirectory = Path.GetDirectoryName(markdownPath) ?? Directory.GetCurrentDirectory();
            }

            if (effectiveOptions.VisualFallbackMode != MarkdownVisualFallbackMode.SvgFile) {
                return effectiveOptions;
            }

            string markdownDirectory = Path.GetDirectoryName(markdownPath) ?? Directory.GetCurrentDirectory();
            if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackDirectory)) {
                string markdownFileName = Path.GetFileNameWithoutExtension(markdownPath);
                string sidecarDirectoryName = string.IsNullOrEmpty(markdownFileName) ? "assets" : markdownFileName + ".assets";
                effectiveOptions.VisualFallbackDirectory = Path.Combine(markdownDirectory, sidecarDirectoryName);
                if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackPathPrefix)) {
                    effectiveOptions.VisualFallbackPathPrefix = sidecarDirectoryName;
                }
                return effectiveOptions;
            }

            if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackPathPrefix)) {
                effectiveOptions.VisualFallbackPathPrefix = MakeMarkdownRelativePath(markdownDirectory, effectiveOptions.VisualFallbackDirectory!);
            }

            return effectiveOptions;
        }

        private static WordToMarkdownOptions CopyOptions(WordToMarkdownOptions source) {
            return new WordToMarkdownOptions {
                FontFamily = source.FontFamily,
                EnableUnderline = source.EnableUnderline,
                EnableHighlight = source.EnableHighlight,
                ImageExportMode = source.ImageExportMode,
                ImageDirectory = source.ImageDirectory,
                FallbackExternalImagesToLinks = source.FallbackExternalImagesToLinks,
                PageBreakMode = source.PageBreakMode,
                UnsupportedContentMode = source.UnsupportedContentMode,
                VisualFallbackMode = source.VisualFallbackMode,
                VisualFallbackDirectory = source.VisualFallbackDirectory,
                VisualFallbackPathPrefix = source.VisualFallbackPathPrefix,
                OnWarning = source.OnWarning,
                IncludeHeadersAndFootersAsSemanticBlocks = source.IncludeHeadersAndFootersAsSemanticBlocks
            };
        }

        private static string MakeMarkdownRelativePath(string baseDirectory, string targetDirectory) {
            try {
                string baseFullPath = Path.GetFullPath(baseDirectory);
                if (!baseFullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) &&
                    !baseFullPath.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
                    baseFullPath += Path.DirectorySeparatorChar;
                }

                string targetFullPath = Path.GetFullPath(targetDirectory);
                var baseUri = new Uri(baseFullPath);
                var targetUri = new Uri(targetFullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                    ? targetFullPath
                    : targetFullPath + Path.DirectorySeparatorChar);
                string relative = Uri.UnescapeDataString(baseUri.MakeRelativeUri(targetUri).ToString());
                return relative.TrimEnd('/').Replace('\\', '/');
            } catch {
                return Path.GetFileName(targetDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            }
        }

        /// <summary>
        /// Converts the document to a Markdown string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>Markdown representation of the document.</returns>
        public static string ToMarkdown(this WordDocument document, WordToMarkdownOptions? options = null) {
            return document.ToMarkdownAsync(options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts the document to a Markdown string asynchronously.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>Markdown representation of the document.</returns>
        public static async Task<string> ToMarkdownAsync(this WordDocument document, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            var converter = new WordToMarkdownConverter();
            var markdown = await converter.ConvertToDocumentAsync(document, options ?? new WordToMarkdownOptions(), cancellationToken).ConfigureAwait(false);
            return markdown.ToMarkdown().Replace("\r\n", "\n").TrimEnd('\n');
        }

        /// <summary>
        /// Converts the document into a typed <see cref="MarkdownDoc"/> without flattening through markdown text first.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>Typed markdown document.</returns>
        public static MarkdownDoc ToMarkdownDocument(this WordDocument document, WordToMarkdownOptions? options = null) {
            return document.ToMarkdownDocumentAsync(options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts the document into a typed <see cref="MarkdownDoc"/> without flattening through markdown text first.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>Typed markdown document.</returns>
        public static Task<MarkdownDoc> ToMarkdownDocumentAsync(this WordDocument document, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            var converter = new WordToMarkdownConverter();
            return converter.ConvertToDocumentAsync(document, options ?? new WordToMarkdownOptions(), cancellationToken);
        }

        /// <summary>
        /// Creates a new document from a Markdown string.
        /// </summary>
        /// <param name="markdown">Markdown content to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromMarkdown(this string markdown, MarkdownToWordOptions? options = null) {
            var converter = new MarkdownToWordConverter();
            return converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), CancellationToken.None).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates a Word document from Markdown by inserting the generated content into an existing template document.
        /// The template retains its styles, sections, headers, footers, table of contents and other package settings.
        /// </summary>
        /// <param name="markdown">Markdown content to insert.</param>
        /// <param name="templatePath">Path to the Word template document to copy and populate.</param>
        /// <param name="options">Template insertion options.</param>
        /// <returns>A populated <see cref="WordDocument"/> instance backed by an in-memory copy of the template.</returns>
        public static WordDocument LoadFromMarkdownTemplate(this string markdown, string templatePath, MarkdownToWordTemplateOptions? options = null) {
            if (string.IsNullOrWhiteSpace(templatePath)) {
                throw new ArgumentException("Template path cannot be null or empty.", nameof(templatePath));
            }

            using var file = File.OpenRead(templatePath);
            var stream = new MemoryStream();
            file.CopyTo(stream);
            stream.Position = 0;
            var document = WordDocument.Load(stream);
            var converter = new MarkdownToWordConverter();
            return converter.ConvertIntoTemplate(markdown, document, options ?? new MarkdownToWordTemplateOptions());
        }

        /// <summary>
        /// Creates a new Word document directly from a typed <see cref="MarkdownDoc"/>.
        /// </summary>
        /// <param name="markdown">Typed markdown document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument ToWordDocument(this MarkdownDoc markdown, MarkdownToWordOptions? options = null) {
            var converter = new MarkdownToWordConverter();
            return converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), CancellationToken.None).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Inserts a typed Markdown document into an existing Word template document.
        /// </summary>
        /// <param name="markdown">Typed Markdown document to insert.</param>
        /// <param name="templateDocument">Template document to populate.</param>
        /// <param name="options">Template insertion options.</param>
        /// <returns>The populated template document.</returns>
        public static WordDocument ToWordDocument(this MarkdownDoc markdown, WordDocument templateDocument, MarkdownToWordTemplateOptions? options = null) {
            var converter = new MarkdownToWordConverter();
            return converter.ConvertIntoTemplate(markdown, templateDocument, options ?? new MarkdownToWordTemplateOptions());
        }

        /// <summary>
        /// Creates a new Word document from HTML by first converting to a typed <see cref="MarkdownDoc"/> and then rendering that AST directly.
        /// </summary>
        /// <param name="html">HTML fragment or document to convert.</param>
        /// <param name="htmlOptions">Optional HTML-to-Markdown conversion options.</param>
        /// <param name="wordOptions">Optional Word conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument ToWordDocumentViaMarkdown(this string html, HtmlToMarkdownOptions? htmlOptions = null, MarkdownToWordOptions? wordOptions = null) {
            var markdown = html.ToMarkdownDocument(htmlOptions);
            return markdown.ToWordDocument(wordOptions);
        }

        /// <summary>
        /// Creates a new Word document from HTML read from a stream by first converting to a typed <see cref="MarkdownDoc"/> and then rendering that AST directly.
        /// </summary>
        /// <param name="htmlStream">Readable stream containing HTML markup.</param>
        /// <param name="htmlOptions">Optional HTML-to-Markdown conversion options.</param>
        /// <param name="wordOptions">Optional Word conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument ToWordDocumentViaMarkdown(this Stream htmlStream, HtmlToMarkdownOptions? htmlOptions = null, MarkdownToWordOptions? wordOptions = null) {
            if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
            using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            return reader.ReadToEnd().ToWordDocumentViaMarkdown(htmlOptions, wordOptions);
        }

        /// <summary>
        /// Creates a new document from a Markdown stream.
        /// </summary>
        /// <param name="markdownStream">Stream containing Markdown content.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromMarkdown(this Stream markdownStream, MarkdownToWordOptions? options = null) {
            return LoadFromMarkdownAsync(markdownStream, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates a new document from a Markdown file.
        /// </summary>
        /// <param name="path">Path to the Markdown file.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="encoding">Encoding to use when reading the file. If <c>null</c>, the encoding is automatically detected from the file's byte order mark.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromMarkdown(string path, MarkdownToWordOptions? options = null, Encoding? encoding = null) {
            using var reader = new StreamReader(path, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: encoding == null);
            string markdown = reader.ReadToEnd();
            var converter = new MarkdownToWordConverter();
            return converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), CancellationToken.None).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously creates a new document from a Markdown string read from the specified path.
        /// </summary>
        /// <param name="path">Path to the Markdown file.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static async Task<WordDocument> LoadFromMarkdownAsync(this string path, MarkdownToWordOptions? options = null, CancellationToken cancellationToken = default) {
#if NET8_0_OR_GREATER
            using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            var markdown = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
            using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            var markdown = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
            var converter = new MarkdownToWordConverter();
            return await converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously creates a new document from a Markdown stream.
        /// </summary>
        /// <param name="markdownStream">Stream containing Markdown content.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static async Task<WordDocument> LoadFromMarkdownAsync(this Stream markdownStream, MarkdownToWordOptions? options = null, CancellationToken cancellationToken = default) {
            using var reader = new StreamReader(markdownStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string markdown = await reader.ReadToEndAsync().ConfigureAwait(false);
            var converter = new MarkdownToWordConverter();
            return await converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously creates a new Word document directly from a typed <see cref="MarkdownDoc"/>.
        /// </summary>
        /// <param name="markdown">Typed markdown document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static Task<WordDocument> ToWordDocumentAsync(this MarkdownDoc markdown, MarkdownToWordOptions? options = null, CancellationToken cancellationToken = default) {
            var converter = new MarkdownToWordConverter();
            return converter.ConvertAsync(markdown, options ?? new MarkdownToWordOptions(), cancellationToken);
        }

        // HTML via OfficeIMO.Markdown (Word -> Markdown -> HTML)

        /// <summary>
        /// Converts the document to a full HTML5 document via OfficeIMO.Markdown.
        /// </summary>
        public static string ToHtmlViaMarkdown(this WordDocument document, HtmlOptions? options = null) {
            options ??= new HtmlOptions { Kind = HtmlKind.Document, Style = HtmlStyle.Word };
            options.Kind = HtmlKind.Document;
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var model = document.ToMarkdownDocument();
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            return model.ToHtmlDocument(options);
        }

        /// <summary>
        /// Converts the document to an embeddable HTML fragment via OfficeIMO.Markdown.
        /// </summary>
        public static string ToHtmlFragmentViaMarkdown(this WordDocument document, HtmlOptions? options = null) {
            options ??= new HtmlOptions { Kind = HtmlKind.Fragment, Style = HtmlStyle.Word };
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var model = document.ToMarkdownDocument();
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            return model.ToHtmlFragment(options);
        }

        /// <summary>
        /// Saves the document as HTML via OfficeIMO.Markdown. Supports external CSS sidecar when configured in <see cref="HtmlOptions"/>.
        /// </summary>
        public static void SaveAsHtmlViaMarkdown(this WordDocument document, string path, HtmlOptions? options = null) {
            options ??= new HtmlOptions { Kind = HtmlKind.Document, Style = HtmlStyle.Word };
            options.Kind = HtmlKind.Document;
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var model = document.ToMarkdownDocument();
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            model.SaveAsHtml(path, options);
        }

        /// <summary>
        /// Asynchronously saves the document as HTML via OfficeIMO.Markdown.
        /// </summary>
        public static async Task SaveAsHtmlViaMarkdownAsync(this WordDocument document, string path, HtmlOptions? options = null, CancellationToken cancellationToken = default) {
            options ??= new HtmlOptions { Kind = HtmlKind.Document, Style = HtmlStyle.Word };
            options.Kind = HtmlKind.Document;
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var model = await document.ToMarkdownDocumentAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            // MarkdownDoc.SaveAsHtml does sync I/O; for now, delegate synchronously to keep surface small.
            model.SaveAsHtml(path, options);
        }
    }
}
