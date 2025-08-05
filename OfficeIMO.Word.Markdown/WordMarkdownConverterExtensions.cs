using OfficeIMO.Word;
using OfficeIMO.Word.Markdown.Converters;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Markdown {
    public static class WordMarkdownConverterExtensions {
        public static void SaveAsMarkdown(this WordDocument document, string path, WordToMarkdownOptions? options = null) {
            var markdown = document.ToMarkdown(options);
            File.WriteAllText(path, markdown, Encoding.UTF8);
        }

        public static void SaveAsMarkdown(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null) {
            var markdown = document.ToMarkdown(options);
            var bytes = Encoding.UTF8.GetBytes(markdown);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static string ToMarkdown(this WordDocument document, WordToMarkdownOptions? options = null) {
            var converter = new WordToMarkdownConverter();
            return converter.Convert(document, options);
        }

        public static WordDocument LoadFromMarkdown(this string markdown, MarkdownToWordOptions? options = null) {
            var converter = new MarkdownToWordConverter();
            return converter.Convert(markdown, options);
        }

        public static WordDocument LoadFromMarkdown(this Stream markdownStream, MarkdownToWordOptions? options = null) {
            using var reader = new StreamReader(markdownStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string markdown = reader.ReadToEnd();
            return LoadFromMarkdown(markdown, options);
        }
    }
}