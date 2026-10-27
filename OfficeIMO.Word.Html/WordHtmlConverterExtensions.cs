using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Extension methods enabling HTML conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static partial class WordHtmlConverterExtensions {
        internal static WordHtmlStyleMissingEventArgs OnStyleMissing(WordHtmlImportOptions options, WordParagraph paragraph, string className) {
            var args = new WordHtmlStyleMissingEventArgs(paragraph, className);
            options.StyleMissingHandler?.Invoke(args);
            return args;
        }







        /// <summary>Converts the document to HTML with the shared structured result contract.</summary>
        internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this WordDocument document, WordHtmlExportOptions? options = null) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            WordHtmlExportOptions resolved = (options ?? new WordHtmlExportOptions()).CloneForConversion();
            var converter = new WordToHtmlConverter();
            string html = converter.Convert(document, resolved);
            return new HtmlTextConversionResult(html, resolved.ConversionReport);
        }





        internal static WordHtmlImportOptions CreateWordOptionsForSharedDocument(HtmlInputTrust trust = HtmlInputTrust.Untrusted) {
            return trust == HtmlInputTrust.Trusted
                ? WordHtmlImportOptions.CreateTrustedDocumentProfile()
                : WordHtmlImportOptions.CreateUntrustedHtmlProfile();
        }

        private static WordHtmlImportOptions ResolveWordOptionsForSharedDocument(
            HtmlConversionDocument document,
            WordHtmlImportOptions? options) {
            WordHtmlImportOptions resolved = (options ?? CreateWordOptionsForSharedDocument(document.Trust)).Clone();
            resolved.HyperlinkUrlPolicy = HtmlUrlPolicy.Intersect(document.HyperlinkUrlPolicy, resolved.HyperlinkUrlPolicy);
            resolved.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(document.ResourceUrlPolicy, resolved.ResourceUrlPolicy);
            resolved.Limits = HtmlConversionLimits.Intersect(document.Limits, resolved.Limits);
            return resolved;
        }

        /// <summary>
        /// Appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToBody(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHtmlImportOptions? options = null) {
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
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
        public static async Task AddHtmlToBodyAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
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
        public static void AddHtmlToHeader(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHeaderFooterType? type = null, WordHtmlImportOptions? options = null) {
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
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
        public static async Task AddHtmlToHeaderAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHeaderFooterType? type = null, WordHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            type ??= WordHeaderFooterType.Default;

            // Prefer section-scoped headers to avoid multi-section warnings
            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the header.");
            doc.AddHeadersAndFooters();
            var headers = targetSection.Header ?? throw new System.InvalidOperationException("The target section does not have any headers defined. Call AddHeadersAndFooters() before appending HTML to the header.");
            var header = GetOrCreateHeader(doc, targetSection, headers, type.Value.ToOpenXml());

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToHeaderAsync(doc, targetSection, header, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToFooter(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHeaderFooterType? type = null, WordHtmlImportOptions? options = null) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));

            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            EnsureOfflineSynchronousImport(htmlDocument, resolved);

            var footerType = type ?? WordHeaderFooterType.Default;
            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the footer.");
            doc.AddHeadersAndFooters();
            var footers = targetSection.Footer ?? throw new System.InvalidOperationException("The target section does not have any footers defined. Call AddHeadersAndFooters() before appending HTML to the footer.");
            GetOrCreateFooter(doc, targetSection, footers, footerType.ToOpenXml());

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
        public static async Task AddHtmlToFooterAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, WordHeaderFooterType? type = null, WordHtmlImportOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            WordHtmlImportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            var footerType = type ?? WordHeaderFooterType.Default;

            var targetSection = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the footer.");
            doc.AddHeadersAndFooters();
            var footers = targetSection.Footer ?? throw new System.InvalidOperationException("The target section does not have any footers defined. Call AddHeadersAndFooters() before appending HTML to the footer.");
            var footer = GetOrCreateFooter(doc, targetSection, footers, footerType.ToOpenXml());

            var converter = new HtmlToWordConverter();
            await converter.AddHtmlToFooterAsync(doc, targetSection, footer, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
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




    }
}
