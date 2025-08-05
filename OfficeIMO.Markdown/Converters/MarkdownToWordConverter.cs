using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Markdown {
    /// <summary>
    /// Converts Markdown text into a Word document without intermediate formats.
    /// </summary>
    public class MarkdownToWordConverter : IWordConverter {
        private static readonly Regex _imageRegex = new("!\\[(.*?)\\]\\((.*?)\\)");
        /// <summary>
        /// Converts Markdown content to DOCX and writes it to the provided stream.
        /// </summary>
        /// <param name="markdown">Markdown source text.</param>
        /// <param name="output">Destination stream for the generated document.</param>
        /// <param name="options">Optional conversion settings.</param>
        public static void Convert(string markdown, Stream output, MarkdownToWordOptions? options = null) {
            if (markdown == null) {
                throw new ConversionException($"{nameof(markdown)} cannot be null.");
            }
            if (output == null) {
                throw new ConversionException($"{nameof(output)} cannot be null.");
            }

            options ??= new MarkdownToWordOptions();
            var fontFamily = FontResolver.Resolve(options.FontFamily);

            using var document = WordDocument.Create();
            WordList? currentList = null;
            bool currentListIsNumbered = false;

            foreach (var raw in markdown.Replace("\r", string.Empty).Split('\n')) {
                var line = raw.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) {
                    currentList = null;
                    continue;
                }

                if (line.StartsWith("#")) {
                    currentList = null;
                    int level = line.TakeWhile(c => c == '#').Count();
                    string text = line.Substring(level).TrimStart();
                    var paragraph = document.AddParagraph();
                    InlineRunHelper.AddInlineRuns(paragraph, text, fontFamily);
                    paragraph.Style = level switch {
                        1 => WordParagraphStyles.Heading1,
                        2 => WordParagraphStyles.Heading2,
                        3 => WordParagraphStyles.Heading3,
                        4 => WordParagraphStyles.Heading4,
                        5 => WordParagraphStyles.Heading5,
                        6 => WordParagraphStyles.Heading6,
                        _ => WordParagraphStyles.Normal
                    };
                    continue;
                }

                var numberMatch = Regex.Match(line, @"^(\d+)\.\s+(.*)");
                if (line.StartsWith("- ") || line.StartsWith("* ") || numberMatch.Success) {
                    bool isNumbered = numberMatch.Success;
                    string text = isNumbered ? numberMatch.Groups[2].Value : line.Substring(2).TrimStart();

                    if (currentList == null || currentListIsNumbered != isNumbered) {
                        currentList = document.AddList(isNumbered ? WordListStyle.ArticleSections : WordListStyle.Bulleted);
                        currentListIsNumbered = isNumbered;
                    }

                    var item = currentList.AddItem(string.Empty);
                    AddInlineContent(item, text, fontFamily);
                    continue;
                }

                currentList = null;
                var para = document.AddParagraph();
                AddInlineContent(para, line, fontFamily);
            }

            document.Save(output);
        }

        private static void AddInlineContent(WordParagraph paragraph, string text, string? fontFamily) {
            int position = 0;
            foreach (Match match in _imageRegex.Matches(text)) {
                string before = text.Substring(position, match.Index - position);
                if (!string.IsNullOrEmpty(before)) {
                    InlineRunHelper.AddInlineRuns(paragraph, before, fontFamily);
                }

                string alt = match.Groups[1].Value;
                string src = match.Groups[2].Value;
                EmbedImage(paragraph, src, alt);

                position = match.Index + match.Length;
            }

            string after = text.Substring(position);
            if (!string.IsNullOrEmpty(after)) {
                InlineRunHelper.AddInlineRuns(paragraph, after, fontFamily);
            }
        }

        private static void EmbedImage(WordParagraph paragraph, string src, string alt) {
            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                int commaIndex = src.IndexOf(',');
                string base64Data = src.Substring(commaIndex + 1);
                int semicolonIndex = src.IndexOf(';');
                string extension = ".png";
                if (semicolonIndex > 5) {
                    string mime = src.Substring(5, semicolonIndex - 5);
                    extension = mime switch {
                        "image/jpeg" => ".jpg",
                        "image/jpg" => ".jpg",
                        "image/gif" => ".gif",
                        "image/bmp" => ".bmp",
                        "image/tiff" => ".tiff",
                        _ => ".png"
                    };
                }
                paragraph.AddImageFromBase64(base64Data, "image" + extension, description: alt);
                return;
            }

            if (Uri.TryCreate(src, UriKind.Absolute, out Uri uri)) {
                if (uri.Scheme == Uri.UriSchemeFile) {
                    paragraph.AddImage(uri.LocalPath, description: alt);
                    return;
                }
                if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) {
                    using HttpClient client = new HttpClient();
                    byte[] bytes = client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                    using MemoryStream ms = new MemoryStream(bytes);
                    string fileName = Path.GetFileName(uri.AbsolutePath);
                    if (string.IsNullOrEmpty(fileName)) {
                        fileName = "image";
                    }
                    paragraph.AddImage(ms, fileName, null, null, description: alt);
                    return;
                }
            }

            if (File.Exists(src)) {
                paragraph.AddImage(src, description: alt);
                return;
            }

            throw new InvalidOperationException("Unable to resolve image source: " + src);
        }
        
        public void Convert(Stream input, Stream output, IConversionOptions options) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string markdown = reader.ReadToEnd();
            Convert(markdown, output, options as MarkdownToWordOptions);
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string markdown = await reader.ReadToEndAsync().ConfigureAwait(false);
            Convert(markdown, output, options as MarkdownToWordOptions);
            await output.FlushAsync().ConfigureAwait(false);
        }
    }
}
