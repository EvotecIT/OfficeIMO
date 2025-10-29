using System.Collections.Generic;
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
                var elements = section.Elements;
                if (elements == null || elements.Count == 0) {
                    // Fallback: compose from known collections when Elements isn't available
                    elements = new List<WordElement>(section.Paragraphs.Count + section.Tables.Count);
                    elements.AddRange(section.Paragraphs);
                    elements.AddRange(section.Tables);
                }
                for (int i = 0; i < elements.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var el = elements[i];
                    if (el is WordParagraph p) {
                        // Elements may include multiple WordParagraph wrappers for a single
                        // underlying OpenXml paragraph (one per run/hyperlink). Rendering each
                        // would duplicate content. Only render once per paragraph by processing
                        // the first run wrapper (IsFirstRun), or paragraphs that have no runs at all.
                        bool hasRuns = false;
                        try { hasRuns = p.GetRuns().Any(); } catch { /* best-effort */ }
                        // Detect checkbox state across sibling wrappers for the same underlying paragraph
                        bool paraHasCheckbox = p.IsCheckBox;
                        bool paraCheckboxChecked = p.CheckBox?.IsChecked == true;
                        // Look ahead within the same paragraph group
                        int j = i + 1;
                        while (j < elements.Count && elements[j] is WordParagraph p2 && p2.Equals(p)) {
                            if (p2.IsCheckBox) { paraHasCheckbox = true; paraCheckboxChecked = p2.CheckBox?.IsChecked == true; }
                            j++;
                        }
                        if (hasRuns && !p.IsFirstRun) {
                            continue; // skip subsequent run wrappers
                        }
                        var text = ConvertParagraph(p, options, paraHasCheckbox, paraCheckboxChecked);
                        if (!string.IsNullOrEmpty(text)) {
                            _output.AppendLine(text);
                            // Ensure a blank line separator between standalone paragraphs so
                            // Markdown renderers don’t merge them into a single paragraph.
                            if (!p.IsListItem) {
                                _output.AppendLine();
                            }
                        }
                    } else if (el is WordTable t) {
                        var tableText = ConvertTable(t, options);
                        if (!string.IsNullOrEmpty(tableText)) {
                            _output.AppendLine(tableText);
                            _output.AppendLine(); // ensure separation so following content isn't merged into table
                        }
                    } else if (el is WordEmbeddedDocument ed) {
                        var html = ed.GetHtml();
                        if (!string.IsNullOrEmpty(html)) {
                            _output.AppendLine(html);
                        }
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
