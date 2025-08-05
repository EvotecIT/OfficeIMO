using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Converters;

namespace OfficeIMO.Markdown {
    /// <summary>
    /// Converts Markdown text into a Word document without intermediate formats.
    /// </summary>
    public class MarkdownToWordConverter : IWordConverter {
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
                    InlineRunHelper.AddInlineRuns(item, text, fontFamily);
                    continue;
                }

                currentList = null;
                var para = document.AddParagraph();
                InlineRunHelper.AddInlineRuns(para, line, fontFamily);
            }

            document.Save(output);
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
