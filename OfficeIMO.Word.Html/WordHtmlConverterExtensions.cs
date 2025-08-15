using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Converters;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Extension methods enabling HTML conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static class WordHtmlConverterExtensions {
        public static event EventHandler<StyleMissingEventArgs>? StyleMissing;

        internal static StyleMissingEventArgs OnStyleMissing(WordParagraph paragraph, string className) {
            var args = new StyleMissingEventArgs(paragraph, className);
            StyleMissing?.Invoke(null, args);
            return args;
        }

        /// <summary>
        /// Saves the document as an HTML file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, string path, WordToHtmlOptions? options = null) {
            document.SaveAsHtmlAsync(path, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Saves the document as HTML to the provided stream.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, Stream stream, WordToHtmlOptions? options = null) {
            document.SaveAsHtmlAsync(stream, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Converts the document to an HTML string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>HTML representation of the document.</returns>
        public static string ToHtml(this WordDocument document, WordToHtmlOptions? options = null) {
            return document.ToHtmlAsync(options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously converts the document to an HTML string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>HTML representation of the document.</returns>
        public static async Task<string> ToHtmlAsync(this WordDocument document, WordToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            cancellationToken.ThrowIfCancellationRequested();
            var converter = new WordToHtmlConverter();
            return await converter.ConvertAsync(document, options ?? new WordToHtmlOptions(), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Creates a new document from an HTML string.
        /// </summary>
        /// <param name="html">HTML content to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromHtml(this string html, HtmlToWordOptions? options = null) {
            return LoadFromHtmlAsync(html, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously creates a new document from an HTML string.
        /// </summary>
        /// <param name="html">HTML content to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static async Task<WordDocument> LoadFromHtmlAsync(this string html, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (html == null) throw new System.ArgumentNullException(nameof(html));
            cancellationToken.ThrowIfCancellationRequested();
            var converter = new HtmlToWordConverter();
            return await converter.ConvertAsync(html, options ?? new HtmlToWordOptions(), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Creates a new document from an HTML stream.
        /// </summary>
        /// <param name="htmlStream">Stream containing HTML content.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromHtml(this Stream htmlStream, HtmlToWordOptions? options = null) {
            return LoadFromHtmlAsync(htmlStream, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously creates a new document from an HTML stream.
        /// </summary>
        /// <param name="htmlStream">Stream containing HTML content.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static async Task<WordDocument> LoadFromHtmlAsync(this Stream htmlStream, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (htmlStream == null) throw new System.ArgumentNullException(nameof(htmlStream));
            cancellationToken.ThrowIfCancellationRequested();
            using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
#if NET8_0_OR_GREATER
            string html = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
            string html = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
            return await LoadFromHtmlAsync(html, options, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToBody(this WordDocument doc, string html, HtmlToWordOptions? options = null) {
            doc.AddHtmlToBodyAsync(html, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToBodyAsync(this WordDocument doc, string html, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (html == null) throw new System.ArgumentNullException(nameof(html));
            cancellationToken.ThrowIfCancellationRequested();

            options ??= new HtmlToWordOptions();

            var section = doc.Sections.Last();
            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToBodyAsync(doc, section, html, options, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Appends HTML content to the document's header.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Header type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToHeader(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            doc.AddHtmlToHeaderAsync(html, type, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's header.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Header type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToHeaderAsync(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (html == null) throw new System.ArgumentNullException(nameof(html));
            cancellationToken.ThrowIfCancellationRequested();

            doc.AddHeadersAndFooters();
            options ??= new HtmlToWordOptions();
            type ??= HeaderFooterValues.Default;

            WordHeader header;
            if (type == HeaderFooterValues.First) {
                header = doc.Header.First;
            } else if (type == HeaderFooterValues.Even) {
                header = doc.Header.Even;
            } else {
                header = doc.Header.Default;
            }

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToHeaderAsync(doc, header, html, options, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToFooter(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            doc.AddHtmlToFooterAsync(html, type, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToFooterAsync(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (html == null) throw new System.ArgumentNullException(nameof(html));
            cancellationToken.ThrowIfCancellationRequested();

            doc.AddHeadersAndFooters();
            options ??= new HtmlToWordOptions();
            type ??= HeaderFooterValues.Default;

            WordFooter footer;
            if (type == HeaderFooterValues.First) {
                footer = doc.Footer.First;
            } else if (type == HeaderFooterValues.Even) {
                footer = doc.Footer.Even;
            } else {
                footer = doc.Footer.Default;
            }

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToFooterAsync(doc, footer, html, options, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Asynchronously saves the document as an HTML file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task SaveAsHtmlAsync(this WordDocument document, string path, WordToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            if (path == null) throw new System.ArgumentNullException(nameof(path));
            cancellationToken.ThrowIfCancellationRequested();
            var html = await document.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_0_OR_GREATER
            await File.WriteAllTextAsync(path, html, Encoding.UTF8, cancellationToken).ConfigureAwait(false);
#else
            using var writer = new StreamWriter(path, false, Encoding.UTF8);
            await writer.WriteAsync(html).ConfigureAwait(false);
#endif
        }

        /// <summary>
        /// Asynchronously saves the document as HTML to the provided stream.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task SaveAsHtmlAsync(this WordDocument document, Stream stream, WordToHtmlOptions? options = null, CancellationToken cancellationToken = default) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            if (stream == null) throw new System.ArgumentNullException(nameof(stream));
            cancellationToken.ThrowIfCancellationRequested();
            var html = await document.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
            var bytes = Encoding.UTF8.GetBytes(html);
#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP3_0_OR_GREATER
            await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
#else
            await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
        }
    }
}