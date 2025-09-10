using OfficeIMO.Word;
using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

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
        private readonly StringBuilder _output = new StringBuilder();

        public string Convert(WordDocument document, WordToMarkdownOptions options) {
            return ConvertAsync(document, options).GetAwaiter().GetResult();
        }

        public Task<string> ConvertAsync(WordDocument document, WordToMarkdownOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToMarkdownOptions();

            _output.Clear();
            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var paragraph in section.Paragraphs) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var text = ConvertParagraph(paragraph, options);
                    if (!string.IsNullOrEmpty(text)) {
                        _output.AppendLine(text);
                    }
                }

                foreach (var table in section.Tables) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var tableText = ConvertTable(table, options);
                    if (!string.IsNullOrEmpty(tableText)) {
                        _output.AppendLine(tableText);
                    }
                }

                foreach (var embedded in section.EmbeddedDocuments) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var html = embedded.GetHtml();
                    if (!string.IsNullOrEmpty(html)) {
                        _output.AppendLine(html);
                    }
                }
            }

            if (document.FootNotes.Count > 0) {
                _output.AppendLine();
                foreach (var footnote in document.FootNotes.OrderBy(fn => fn.ReferenceId)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (footnote.ReferenceId.HasValue) {
                        _output.AppendLine($"[^{footnote.ReferenceId}]: {RenderFootnote(footnote, options)}");
                    }
                }
            }

            return Task.FromResult(_output.ToString().TrimEnd());
        }
    }
}
