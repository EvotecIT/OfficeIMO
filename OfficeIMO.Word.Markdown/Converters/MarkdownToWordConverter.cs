using Markdig;
using Markdig.Extensions.Footnotes;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown.Converters {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Use Markdig to parse markdown into AST (Abstract Syntax Tree)
    /// 2. Convert Markdig elements to OfficeIMO.Word API calls:
    ///    - HeadingBlock -> wordDoc.AddParagraph(text).Style = WordParagraphStyles.Heading1/2/3...
    ///    - ListBlock -> wordDoc.AddList() with appropriate style
    ///    - CodeBlock -> wordDoc.AddParagraph() with monospace font
    ///    - Table -> wordDoc.AddTable()
    /// 3. For inline formatting:
    ///    - EmphasisInline (single) -> paragraph.AddText(text).Italic = true
    ///    - EmphasisInline (double) -> paragraph.AddText(text).Bold = true
    ///    - LinkInline -> paragraph.AddHyperLink()
    /// 4. Reuse existing OfficeIMO.Word functionality, don't recreate
    /// </summary>
    internal partial class MarkdownToWordConverter {
        public WordDocument Convert(string markdown, MarkdownToWordOptions options) {
            return ConvertAsync(markdown, options).GetAwaiter().GetResult();
        }

        public Task<WordDocument> ConvertAsync(string markdown, MarkdownToWordOptions options, CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);

            var pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .UseFootnotes()
                .Build();
            MarkdownDocument markdownDocument = Markdig.Markdown.Parse(markdown, pipeline);

            foreach (var block in markdownDocument) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessBlock(block, document, options);
            }

            return Task.FromResult(document);
        }

        private static void ProcessBlock(Block block, WordDocument document, MarkdownToWordOptions options, WordList? currentList = null, int listLevel = 0, int quoteDepth = 0) {
            switch (block) {
                case HeadingBlock heading:
                    var headingParagraph = document.AddParagraph(string.Empty);
                    if (quoteDepth > 0) {
                        headingParagraph.IndentationBefore = 720 * quoteDepth;
                    }
                    ProcessInline(heading.Inline, headingParagraph, options, document);
                    headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(heading.Level);
                    break;
                case HtmlBlock htmlBlock:
                    string html = htmlBlock.Lines.ToString();
                    document.AddHtmlToBody(html);
                    break;
                case ParagraphBlock paragraphBlock when currentList == null:
                    var paragraph = document.AddParagraph(string.Empty);
                    if (quoteDepth > 0) {
                        paragraph.IndentationBefore = 720 * quoteDepth;
                    }
                    ProcessInline(paragraphBlock.Inline, paragraph, options, document);
                    break;
                case ParagraphBlock paragraphBlock:
                    var listItemParagraph = currentList!.AddItem(string.Empty, listLevel);
                    if (quoteDepth > 0) {
                        listItemParagraph.IndentationBefore = 720 * quoteDepth;
                    }
                    ProcessInline(paragraphBlock.Inline, listItemParagraph, options, document);
                    break;
                case ListBlock listBlock:
                    ProcessListBlock(listBlock, document, options, currentList, listLevel);
                    break;
                case QuoteBlock quote:
                    foreach (var sub in quote) {
                        ProcessBlock(sub, document, options, null, 0, quoteDepth + 1);
                    }
                    break;
                case CodeBlock codeBlock:
                    var codeParagraph = document.AddParagraph(string.Empty);
                    if (quoteDepth > 0) {
                        codeParagraph.IndentationBefore = 720 * quoteDepth;
                    }
                    var codeText = GetCodeBlockText(codeBlock);
                    var run = codeParagraph.AddFormattedText(codeText);
                    var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                    run.SetFontFamily(monoFont);
                    if (codeBlock is FencedCodeBlock fenced && !string.IsNullOrWhiteSpace(fenced.Info)) {
                        var info = (fenced.Info ?? string.Empty).Split(new[] { ' ', '{' }, StringSplitOptions.RemoveEmptyEntries);
                        if (info.Length > 0) {
                            var lang = new string(info[0].Where(char.IsLetterOrDigit).ToArray());
                            if (!string.IsNullOrEmpty(lang)) {
                                codeParagraph.SetStyleId($"CodeLang_{lang}");
                            }
                        }
                    }
                    break;
                case Table table:
                    ProcessTable(table, document, options);
                    break;
                case ThematicBreakBlock:
                    document.AddHorizontalLine();
                    break;
                case FootnoteGroup:
                    // Footnote definitions are processed when their links are encountered
                    break;
            }
        }
    }
}