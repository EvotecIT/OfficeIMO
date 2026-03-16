using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Read document content using OfficeIMO.Word API:
    ///    - document.Paragraphs for text content
    ///    - paragraph.Style to determine heading levels
    ///    - document.Lists for bullet/numbered lists
    ///    - document.Tables for tables
    /// 2. Convert OfficeIMO.Word elements to Markdown syntax:
    ///    - WordParagraphStyles.Heading1 -> # Heading
    ///    - WordParagraphStyles.Heading2 -> ## Heading
    ///    - Bold text -> **text**
    ///    - Italic text -> *text*
    ///    - Lists -> - item or 1. item
    ///    - Tables -> | col1 | col2 |
    /// 3. Check paragraph.IsListItem to identify list items
    /// 4. Use paragraph.Bold, paragraph.Italic for inline formatting
    /// </summary>
    internal partial class WordToMarkdownConverter {
        public string Convert(WordDocument document, WordToMarkdownOptions options) {
            return ConvertAsync(document, options).GetAwaiter().GetResult();
        }

        public Task<string> ConvertAsync(WordDocument document, WordToMarkdownOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToMarkdownOptions();
            return ConvertToDocumentAsync(document, options, cancellationToken)
                .ContinueWith(
                    static task => NormalizeMarkdownLineEndings(task.GetAwaiter().GetResult().ToMarkdown()),
                    cancellationToken,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
        }

        public Task<MarkdownDoc> ConvertToDocumentAsync(WordDocument document, WordToMarkdownOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToMarkdownOptions();

            var markdown = MarkdownDoc.Create();
            BuildMarkdownDocument(document, markdown, options, cancellationToken);
            return Task.FromResult(markdown);
        }

        private static string NormalizeMarkdownLineEndings(string markdown) {
            return string.IsNullOrEmpty(markdown)
                ? string.Empty
                : markdown.Replace("\r\n", "\n").TrimEnd('\n');
        }
    }
}
