using OfficeIMO.Word;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Html {
    public static class WordHtmlConverterExtensions {
        public static void SaveAsHtml(this WordDocument document, string path, WordToHtmlOptions? options = null) {
            using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
            SaveAsHtml(document, stream, options);
        }

        public static void SaveAsHtml(this WordDocument document, Stream stream, WordToHtmlOptions? options = null) {
            using var tempStream = new MemoryStream();
            document.Save(tempStream);
            tempStream.Position = 0;

            string html = WordToHtmlConverter.Convert(tempStream, options);
            byte[] bytes = Encoding.UTF8.GetBytes(html);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static string ToHtml(this WordDocument document, WordToHtmlOptions? options = null) {
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            return WordToHtmlConverter.Convert(stream, options);
        }

        public static WordDocument LoadFromHtml(string html, HtmlToWordOptions? options = null) {
            using var stream = new MemoryStream();
            HtmlToWordConverter.Convert(html, stream, options);
            stream.Position = 0;
            return WordDocument.Load(stream);
        }

        public static WordDocument LoadFromHtml(Stream htmlStream, HtmlToWordOptions? options = null) {
            using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string html = reader.ReadToEnd();
            return LoadFromHtml(html, options);
        }
    }
}