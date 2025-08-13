using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Converters;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Extension methods enabling HTML conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static class WordHtmlConverterExtensions {
        /// <summary>
        /// Saves the document as an HTML file at the specified path.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, string path, WordToHtmlOptions? options = null) {
            var html = document.ToHtml(options);
            File.WriteAllText(path, html, Encoding.UTF8);
        }

        /// <summary>
        /// Saves the document as HTML to the provided stream.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void SaveAsHtml(this WordDocument document, Stream stream, WordToHtmlOptions? options = null) {
            var html = document.ToHtml(options);
            var bytes = Encoding.UTF8.GetBytes(html);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// Converts the document to an HTML string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>HTML representation of the document.</returns>
        public static string ToHtml(this WordDocument document, WordToHtmlOptions? options = null) {
            var converter = new WordToHtmlConverter();
            // Use GetAwaiter().GetResult() to call async method synchronously
            return converter.ConvertAsync(document, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates a new document from an HTML string.
        /// </summary>
        /// <param name="html">HTML content to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromHtml(this string html, HtmlToWordOptions? options = null) {
            var converter = new HtmlToWordConverter();
            // Use GetAwaiter().GetResult() to call async method synchronously
            return converter.ConvertAsync(html, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates a new document from an HTML stream.
        /// </summary>
        /// <param name="htmlStream">Stream containing HTML content.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadFromHtml(this Stream htmlStream, HtmlToWordOptions? options = null) {
            using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string html = reader.ReadToEnd();
            return LoadFromHtml(html, options);
        }

        /// <summary>
        /// Appends HTML content to the document's header.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Header type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToHeader(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (html == null) throw new System.ArgumentNullException(nameof(html));

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
            converter.AddHtmlToHeaderAsync(doc, header, html, options).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Appends HTML content to the document's footer.
        /// </summary>
        /// <param name="doc">Document to modify.</param>
        /// <param name="html">HTML fragment to insert.</param>
        /// <param name="type">Footer type to target.</param>
        /// <param name="options">Optional conversion options.</param>
        public static void AddHtmlToFooter(this WordDocument doc, string html, HeaderFooterValues? type = null, HtmlToWordOptions? options = null) {
            if (doc == null) throw new System.ArgumentNullException(nameof(doc));
            if (html == null) throw new System.ArgumentNullException(nameof(html));

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
            converter.AddHtmlToFooterAsync(doc, footer, html, options).GetAwaiter().GetResult();
        }
    }
}