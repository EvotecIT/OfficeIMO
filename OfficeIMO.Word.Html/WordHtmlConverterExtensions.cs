using OfficeIMO.Word;
using OfficeIMO.Word.Html.Converters;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Html {
    public static class WordHtmlConverterExtensions {
        public static void SaveAsHtml(this WordDocument document, string path, WordToHtmlOptions? options = null) {
            var html = document.ToHtml(options);
            File.WriteAllText(path, html, Encoding.UTF8);
        }

        public static void SaveAsHtml(this WordDocument document, Stream stream, WordToHtmlOptions? options = null) {
            var html = document.ToHtml(options);
            var bytes = Encoding.UTF8.GetBytes(html);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static string ToHtml(this WordDocument document, WordToHtmlOptions? options = null) {
            var converter = new WordToHtmlConverter();
            // Use GetAwaiter().GetResult() to call async method synchronously
            return converter.ConvertAsync(document, options).GetAwaiter().GetResult();
        }

        public static WordDocument LoadFromHtml(this string html, HtmlToWordOptions? options = null) {
            var converter = new HtmlToWordConverter();
            // Use GetAwaiter().GetResult() to call async method synchronously
            return converter.ConvertAsync(html, options).GetAwaiter().GetResult();
        }

        public static WordDocument LoadFromHtml(this Stream htmlStream, HtmlToWordOptions? options = null) {
            using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string html = reader.ReadToEnd();
            return LoadFromHtml(html, options);
        }
    }
}