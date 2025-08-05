using OfficeIMO.Word;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Markdown {
    public static class WordMarkdownConverterExtensions {
        public static void SaveAsMarkdown(this WordDocument document, string path, WordToMarkdownOptions? options = null) {
            using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
            SaveAsMarkdown(document, stream, options);
        }

        public static void SaveAsMarkdown(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null) {
            using var tempStream = new MemoryStream();
            document.Save(tempStream);
            tempStream.Position = 0;

            string markdown = WordToMarkdownConverter.Convert(tempStream, options);
            byte[] bytes = Encoding.UTF8.GetBytes(markdown);
            stream.Write(bytes, 0, bytes.Length);
        }

        public static string ToMarkdown(this WordDocument document, WordToMarkdownOptions? options = null) {
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            return WordToMarkdownConverter.Convert(stream, options);
        }

        public static WordDocument LoadFromMarkdown(string markdown, MarkdownToWordOptions? options = null) {
            using var stream = new MemoryStream();
            MarkdownToWordConverter.Convert(markdown, stream, options);
            stream.Position = 0;
            return WordDocument.Load(stream);
        }

        public static WordDocument LoadFromMarkdown(Stream markdownStream, MarkdownToWordOptions? options = null) {
            using var reader = new StreamReader(markdownStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            string markdown = reader.ReadToEnd();
            return LoadFromMarkdown(markdown, options);
        }
    }
}