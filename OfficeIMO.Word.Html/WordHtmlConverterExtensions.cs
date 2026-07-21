using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Extension methods enabling HTML conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static partial class WordHtmlConverterExtensions {
        internal static StyleMissingEventArgs OnStyleMissing(HtmlToWordOptions options, WordParagraph paragraph, string className) {
            var args = new StyleMissingEventArgs(paragraph, className);
            options.StyleMissingHandler?.Invoke(args);
            return args;
        }

        /// <summary>
        /// Saves the document as an HTML file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, string path, WordToHtmlOptions? options = null) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            if (path == null) throw new System.ArgumentNullException(nameof(path));
            HtmlTextIO.Write(path, document.ToHtml(options));
        }

        /// <summary>
        /// Saves the document as HTML to the provided stream.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, Stream stream, WordToHtmlOptions? options = null) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            if (stream == null) throw new System.ArgumentNullException(nameof(stream));
            HtmlTextIO.Write(stream, document.ToHtml(options));
        }

        /// <summary>
        /// Converts the document to an HTML string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>HTML representation of the document.</returns>
        public static string ToHtml(this WordDocument document, WordToHtmlOptions? options = null) {
            return document.ToHtmlResult(options).Value;
        }

        /// <summary>Converts the document to HTML with the shared structured result contract.</summary>
        public static HtmlTextConversionResult ToHtmlResult(this WordDocument document, WordToHtmlOptions? options = null) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            var converter = new WordToHtmlConverter();
            return new HtmlTextConversionResult(converter.Convert(document, options ?? new WordToHtmlOptions()));
        }

        /// <summary>
        /// Creates a new document from a shared OfficeIMO HTML conversion document.
        /// </summary>
        /// <param name="document">Shared HTML conversion document.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument ToWordDocument(this HtmlConversionDocument document, HtmlToWordOptions? options = null) {
            return document.ToWordDocumentResult(options).RequireValue();
        }

        /// <summary>
        /// Asynchronously creates a new document from a shared OfficeIMO HTML conversion document.
        /// </summary>
        /// <param name="document">Shared HTML conversion document.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static async Task<WordDocument> ToWordDocumentAsync(this HtmlConversionDocument document, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            HtmlToWordResult result = await document.ToWordDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return result.RequireValue();
        }

        internal static HtmlToWordOptions CreateWordOptionsForSharedDocument(HtmlInputTrust trust = HtmlInputTrust.Untrusted) {
            return trust == HtmlInputTrust.Trusted
                ? HtmlToWordOptions.CreateTrustedDocumentProfile()
                : HtmlToWordOptions.CreateUntrustedHtmlProfile();
        }

        private static HtmlToWordOptions ResolveWordOptionsForSharedDocument(
            HtmlConversionDocument document,
            HtmlToWordOptions? options) {
            HtmlToWordOptions resolved = (options ?? CreateWordOptionsForSharedDocument(document.Trust)).Clone();
            resolved.HyperlinkUrlPolicy = HtmlUrlPolicy.Intersect(document.HyperlinkUrlPolicy, resolved.HyperlinkUrlPolicy);
            resolved.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(document.ResourceUrlPolicy, resolved.ResourceUrlPolicy);
            return resolved;
        }

        /// <summary>
        /// Appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToBody(this WordDocument doc, HtmlConversionDocument htmlDocument, HtmlToWordOptions? options = null) {
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            EnsureOfflineSynchronousImport(htmlDocument, resolved);
            doc.AddHtmlToBodyAsync(htmlDocument, resolved).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToBodyAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);

            var section = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the body.");
            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToBodyAsync(doc, section, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Appends HTML content to the document's header.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="type">Header type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToHeader(this WordDocument doc, HtmlConversionDocument htmlDocument, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            EnsureOfflineSynchronousImport(htmlDocument, resolved);
            doc.AddHtmlToHeaderAsync(htmlDocument, type, resolved).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's header.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="type">Header type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToHeaderAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, HeaderFooterValues? type = null, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            type ??= HeaderFooterValues.Default;

            // Prefer section-scoped headers to avoid multi-section warnings
            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the header.");
            doc.AddHeadersAndFooters();
            var headers = targetSection.Header ?? throw new System.InvalidOperationException("The target section does not have any headers defined. Call AddHeadersAndFooters() before appending HTML to the header.");
            var header = GetOrCreateHeader(doc, targetSection, headers, type.Value);

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToHeaderAsync(doc, header, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToFooter(this WordDocument doc, HtmlConversionDocument htmlDocument, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));

            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            EnsureOfflineSynchronousImport(htmlDocument, resolved);

            var footerType = type ?? HeaderFooterValues.Default;
            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the footer.");
            doc.AddHeadersAndFooters();
            var footers = targetSection.Footer ?? throw new System.InvalidOperationException("The target section does not have any footers defined. Call AddHeadersAndFooters() before appending HTML to the footer.");
            GetOrCreateFooter(doc, targetSection, footers, footerType);

            doc.AddHtmlToFooterAsync(htmlDocument, footerType, resolved).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        public static async Task AddHtmlToFooterAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, HeaderFooterValues? type = null, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            HtmlToWordOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            var footerType = type ?? HeaderFooterValues.Default;

            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the footer.");
            doc.AddHeadersAndFooters();
            var footers = targetSection.Footer ?? throw new System.InvalidOperationException("The target section does not have any footers defined. Call AddHeadersAndFooters() before appending HTML to the footer.");
            var footer = GetOrCreateFooter(doc, targetSection, footers, footerType);

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToFooterAsync(doc, footer, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        private static WordFooter GetOrCreateFooter(WordDocument document, WordSection targetSection, WordFooters footers, HeaderFooterValues footerType) {
            WordFooter? footer = SelectFooter(footers, footerType);
            if (footer == null) {
                WordHeadersAndFooters.AddFooterReference(document, targetSection, footerType);
                footer = SelectFooter(footers, footerType);
            }

            return footer ?? throw new System.InvalidOperationException($"The {DescribeHeaderFooter(footerType)} footer could not be located or created for the current section.");

            static WordFooter? SelectFooter(WordFooters source, HeaderFooterValues type) {
                if (type == HeaderFooterValues.First) return source.First;
                if (type == HeaderFooterValues.Even) return source.Even;
                return source.Default;
            }
        }

        private static WordHeader GetOrCreateHeader(WordDocument document, WordSection targetSection, WordHeaders headers, HeaderFooterValues headerType) {
            WordHeader? header = SelectHeader(headers, headerType);
            if (header == null) {
                WordHeadersAndFooters.AddHeaderReference(document, targetSection, headerType);
                header = SelectHeader(headers, headerType);
            }

            return header ?? throw new System.InvalidOperationException($"The {DescribeHeaderFooter(headerType)} header could not be located or created for the current section.");

            static WordHeader? SelectHeader(WordHeaders source, HeaderFooterValues type) {
                if (type == HeaderFooterValues.First) return source.First;
                if (type == HeaderFooterValues.Even) return source.Even;
                return source.Default;
            }
        }

        private static string DescribeHeaderFooter(HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) return "first-page";
            if (type == HeaderFooterValues.Even) return "even-page";
            return "default";
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
            string html = document.ToHtml(options);
            cancellationToken.ThrowIfCancellationRequested();
            await HtmlTextIO.WriteAsync(path, html, cancellationToken).ConfigureAwait(false);
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
            string html = document.ToHtml(options);
            cancellationToken.ThrowIfCancellationRequested();
            await HtmlTextIO.WriteAsync(stream, html, cancellationToken).ConfigureAwait(false);
        }
    }
}
