using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Extension methods enabling HTML conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static partial class WordHtmlConverterExtensions {
        internal static WordHtmlStyleMissingEventArgs OnStyleMissing(HtmlWordExportOptions options, WordParagraph paragraph, string className) {
            var args = new WordHtmlStyleMissingEventArgs(paragraph, className);
            options.StyleMissingHandler?.Invoke(args);
            return args;
        }







        /// <summary>Converts the document to HTML with the shared structured result contract.</summary>
        internal static HtmlTextConversionResult ConvertToHtmlWithEvidence(this WordDocument document, WordHtmlExportOptions? options = null,
            System.Threading.CancellationToken cancellationToken = default) {
            if (document == null) throw new System.ArgumentNullException(nameof(document));
            WordHtmlExportOptions resolved = (options ?? new WordHtmlExportOptions()).CloneForConversion();
            var converter = new WordToHtmlConverter();
            string html = converter.Convert(document, resolved, cancellationToken);
            return new HtmlTextConversionResult(html, resolved.ConversionReport);
        }





        internal static HtmlWordExportOptions CreateWordOptionsForSharedDocument(HtmlInputTrust trust = HtmlInputTrust.Untrusted) {
            return trust == HtmlInputTrust.Trusted
                ? HtmlWordExportOptions.CreateTrustedDocumentProfile()
                : HtmlWordExportOptions.UntrustedDefaults;
        }

        private static HtmlWordExportOptions ResolveWordOptionsForSharedDocument(
            HtmlConversionDocument document,
            HtmlWordExportOptions? options) {
            HtmlWordExportOptions resolved = (options ?? CreateWordOptionsForSharedDocument(document.Trust)).Clone();
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
        /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
        public static IReadOnlyList<WordElement> AddHtml(this WordDocument doc, HtmlConversionDocument htmlDocument, HtmlWordExportOptions? options = null) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            HtmlWordExportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);
            EnsureOfflineSynchronousImport(htmlDocument, resolved);
            return doc.AddHtmlAsync(htmlDocument, resolved).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously appends HTML content to the document's body.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="htmlDocument">Parsed HTML source to insert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>Live imported elements in document order, using the existing Word element projection.</returns>
        public static async Task<IReadOnlyList<WordElement>> AddHtmlAsync(this WordDocument doc, HtmlConversionDocument htmlDocument, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (htmlDocument == null) throw new System.ArgumentNullException(nameof(htmlDocument));
            cancellationToken.ThrowIfCancellationRequested();

            HtmlWordExportOptions resolved = ResolveWordOptionsForSharedDocument(htmlDocument, options);
            resolved.ConversionReport.AddRange(htmlDocument.Diagnostics);

            var section = doc.Sections.LastOrDefault() ?? throw new System.InvalidOperationException("The document does not contain any sections to append HTML to the body.");
            var converter = new HtmlToWordConverter();
            IReadOnlyList<WordElement> elements = await converter.AddHtmlToBodyAsync(doc, section, CreateWordSourceDocument(htmlDocument, resolved.ConversionReport), resolved, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return elements;
        }

    }
}
