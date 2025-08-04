using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
                throw new ArgumentNullException(nameof(markdown));
            }
            if (output == null) {
                throw new ArgumentNullException(nameof(output));
            }

            options ??= new MarkdownToWordOptions();

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
                    AddInlineRuns(paragraph, text, options);
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
                    AddInlineRuns(item, text, options);
                    continue;
                }

                currentList = null;
                var para = document.AddParagraph();
                AddInlineRuns(para, line, options);
            }

            document.Save(output);
        }

        private static void AddInlineRuns(WordParagraph paragraph, string text, MarkdownToWordOptions options) {
            var regex = new Regex(@"(\*\*[^\*]+\*\*|\*[^\*]+\*|[^\*]+)", RegexOptions.Singleline);
            foreach (Match match in regex.Matches(text)) {
                string token = match.Value;
                bool bold = token.StartsWith("**") && token.EndsWith("**");
                bool italic = !bold && token.StartsWith("*") && token.EndsWith("*");
                string value = bold ? token.Substring(2, token.Length - 4) :
                               italic ? token.Substring(1, token.Length - 2) : token;

                var run = paragraph.AddText(value);
                if (!string.IsNullOrEmpty(options.FontFamily)) {
                    run.SetFontFamily(options.FontFamily);
                }
                if (bold) {
                    run.SetBold();
                }
                if (italic) {
                    run.SetItalic();
                }
            }
        }
        public void Convert(Stream input, Stream output, IConversionOptions options) {
            if (input == null) {
                throw new ArgumentNullException(nameof(input));
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
    }
}
