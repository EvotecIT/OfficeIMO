using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Markdown;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Extension methods enabling Markdown conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static class WordMarkdownConverterExtensions {
        /// <summary>
        /// Saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsMarkdown(this WordDocument document, string path, WordToMarkdownOptions? options = null) {
            document.SaveAsMarkdownAsync(path, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Saves the document as Markdown to the provided stream.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsMarkdown(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null) {
            document.SaveAsMarkdownAsync(stream, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task SaveAsMarkdownAsync(this WordDocument document, string path, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            options ??= new WordToMarkdownOptions();
            if (options.ImageExportMode == ImageExportMode.File && string.IsNullOrEmpty(options.ImageDirectory)) {
                options.ImageDirectory = Path.GetDirectoryName(path) ?? Directory.GetCurrentDirectory();
            }
            var markdown = await document.ToMarkdownAsync(options, cancellationToken).ConfigureAwait(false);
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
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
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
            return await converter.ConvertAsync(document, options ?? new WordToMarkdownOptions(), cancellationToken).ConfigureAwait(false);
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
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_0_OR_GREATER
            var markdown = await File.ReadAllTextAsync(path, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
#else
            using var reader = new StreamReader(path, Encoding.UTF8);
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

        // HTML via OfficeIMO.Markdown (Word -> Markdown -> HTML)

        /// <summary>
        /// Converts the document to a full HTML5 document via OfficeIMO.Markdown.
        /// </summary>
        public static string ToHtmlViaMarkdown(this WordDocument document, HtmlOptions? options = null) {
            options ??= new HtmlOptions { Kind = HtmlKind.Document, Style = HtmlStyle.Word };
            options.Kind = HtmlKind.Document;
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var md = document.ToMarkdown();
            var model = MarkdownReader.Parse(md);
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
            var md = document.ToMarkdown();
            var model = MarkdownReader.Parse(md);
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
            var md = document.ToMarkdown();
            var model = MarkdownReader.Parse(md);
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            model.SaveHtml(path, options);
        }

        /// <summary>
        /// Asynchronously saves the document as HTML via OfficeIMO.Markdown.
        /// </summary>
        public static async Task SaveAsHtmlViaMarkdownAsync(this WordDocument document, string path, HtmlOptions? options = null, CancellationToken cancellationToken = default) {
            options ??= new HtmlOptions { Kind = HtmlKind.Document, Style = HtmlStyle.Word };
            options.Kind = HtmlKind.Document;
            if (options.Style == default) options.Style = HtmlStyle.Word;
            var md = await document.ToMarkdownAsync().ConfigureAwait(false);
            var model = MarkdownReader.Parse(md);
            if (options.InjectTocAtTop && !model.Blocks.Any(b => string.Equals(b.GetType().Name, "TocPlaceholderBlock", System.StringComparison.Ordinal))) {
                model.TocAtTop(options.InjectTocTitle, options.InjectTocMinLevel, options.InjectTocMaxLevel, options.InjectTocOrdered, options.InjectTocTitleLevel);
            }
            // MarkdownDoc.SaveHtml does sync I/O; for now, delegate synchronously to keep surface small
            model.SaveHtml(path, options);
        }
    }
}
