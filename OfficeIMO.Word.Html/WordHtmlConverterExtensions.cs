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
    }
}