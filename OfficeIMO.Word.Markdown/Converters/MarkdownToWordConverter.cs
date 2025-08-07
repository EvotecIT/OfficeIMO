using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Extensions.Footnotes;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;

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
        private readonly Dictionary<string, string> _footnotes = new();

        public WordDocument Convert(string markdown, MarkdownToWordOptions options) {
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
            var parsed = Markdig.Markdown.Parse(markdown, pipeline);

            _footnotes.Clear();
            var footnoteGroup = parsed.Descendants<FootnoteGroup>();
            foreach (var group in footnoteGroup) {
                foreach (Footnote footnote in group) {
                    var text = new System.Text.StringBuilder();
                    foreach (var fnBlock in footnote) {
                        if (fnBlock is ParagraphBlock pb) {
                            if (text.Length > 0) {
                                text.Append(' ');
                            }
                            text.Append(BuildMarkdown(pb.Inline));
                        }
                    }
                    var label = footnote.Label ?? footnote.Order.ToString();
                    _footnotes[label] = text.ToString();
                }
            }

            foreach (var block in parsed) {
                ProcessBlock(block, document, options, _footnotes);
            }

            return document;
        }

        private static void ProcessBlock(Block block, WordDocument document, MarkdownToWordOptions options, Dictionary<string, string> footnotes, WordList? currentList = null, int listLevel = 0) {
            switch (block) {
                case HeadingBlock heading:
                    var headingParagraph = document.AddParagraph(string.Empty);
                    ProcessInline(heading.Inline, headingParagraph, options, document, footnotes);
                    headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(heading.Level);
                    break;
                case ParagraphBlock paragraphBlock when currentList == null:
                    var paragraph = document.AddParagraph(string.Empty);
                    ProcessInline(paragraphBlock.Inline, paragraph, options, document, footnotes);
                    break;
                case ParagraphBlock paragraphBlock:
                    var listItemParagraph = currentList!.AddItem(string.Empty, listLevel);
                    ProcessInline(paragraphBlock.Inline, listItemParagraph, options, document, footnotes);
                    break;
                case ListBlock listBlock:
                    ProcessListBlock(listBlock, document, options, footnotes, currentList, listLevel);
                    break;
                case QuoteBlock quote:
                    foreach (var sub in quote) {
                        if (sub is ParagraphBlock qp) {
                            var qpParagraph = document.AddParagraph(string.Empty);
                            qpParagraph.IndentationBefore = 720;
                            ProcessInline(qp.Inline, qpParagraph, options, document, footnotes);
                        } else {
                            ProcessBlock(sub, document, options, footnotes);
                        }
                    }
                    break;
                case CodeBlock codeBlock:
                    var codeParagraph = document.AddParagraph(string.Empty);
                    var codeText = GetCodeBlockText(codeBlock);
                    var run = codeParagraph.AddFormattedText(codeText);
                    run.SetFontFamily("Consolas");
                    break;
                case Table table:
                    ProcessTable(table, document, options, footnotes);
                    break;
                case ThematicBreakBlock:
                    document.AddHorizontalLine();
                    break;
                case FootnoteGroup:
                    // Footnotes are processed separately via inline links
                    break;
            }
        }
    }
}