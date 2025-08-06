using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using System.Text;
using Markdig.Extensions.Tables;

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
    internal class MarkdownToWordConverter {
        public WordDocument Convert(string markdown, MarkdownToWordOptions options) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);

            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            var parsed = Markdig.Markdown.Parse(markdown, pipeline);

            foreach (var block in parsed) {
                ProcessBlock(block, document, options);
            }

            return document;
        }

        private static void ProcessBlock(Block block, WordDocument document, MarkdownToWordOptions options, WordList? currentList = null, int listLevel = 0) {
            switch (block) {
                case HeadingBlock heading:
                    var headingParagraph = document.AddParagraph(string.Empty);
                    ProcessInline(heading.Inline, headingParagraph, options, document);
                    headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(heading.Level);
                    break;
                case ParagraphBlock paragraphBlock when currentList == null:
                    var paragraph = document.AddParagraph(string.Empty);
                    ProcessInline(paragraphBlock.Inline, paragraph, options, document);
                    break;
                case ParagraphBlock paragraphBlock:
                    var listItemParagraph = currentList!.AddItem(string.Empty, listLevel);
                    ProcessInline(paragraphBlock.Inline, listItemParagraph, options, document);
                    break;
                case ListBlock listBlock:
                    var list = listBlock.IsOrdered ? document.CreateNumberedList() : document.CreateBulletList();
                    foreach (ListItemBlock listItem in listBlock) {
                        var firstParagraph = listItem.FirstOrDefault() as ParagraphBlock;
                        if (firstParagraph != null) {
                            var listParagraph = list.AddItem(string.Empty, listLevel);
                            ProcessInline(firstParagraph.Inline, listParagraph, options, document);
                        }
                        foreach (var sub in listItem.Skip(1)) {
                            ProcessBlock(sub, document, options, list, listLevel + 1);
                        }
                    }
                    break;
                case QuoteBlock quote:
                    foreach (var sub in quote) {
                        if (sub is ParagraphBlock qp) {
                            var qpParagraph = document.AddParagraph(string.Empty);
                            qpParagraph.IndentationBefore = 720;
                            ProcessInline(qp.Inline, qpParagraph, options, document);
                        } else {
                            ProcessBlock(sub, document, options);
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
                    ProcessTable(table, document, options);
                    break;
                case ThematicBreakBlock:
                    document.AddHorizontalLine();
                    break;
            }
        }

        private static void ProcessTable(Table table, WordDocument document, MarkdownToWordOptions options) {
            int rows = table.Count();
            int cols = table.ColumnDefinitions.Count;
            var wordTable = document.AddTable(rows, cols);
            int r = 0;
            foreach (TableRow row in table) {
                int c = 0;
                foreach (TableCell cell in row) {
                    var target = wordTable.Rows[r].Cells[c].Paragraphs[0];
                    foreach (var cellBlock in cell) {
                        if (cellBlock is ParagraphBlock pb) {
                            ProcessInline(pb.Inline, target, options, document);
                        }
                    }
                    c++;
                }
                r++;
            }
        }

        private static void ProcessInline(Inline? inline, WordParagraph paragraph, MarkdownToWordOptions options, WordDocument document) {
            if (inline == null) {
                return;
            }

            var buffer = new StringBuilder();

            void Flush() {
                if (buffer.Length > 0) {
                    InlineRunHelper.AddInlineRuns(paragraph, buffer.ToString(), options.FontFamily);
                    buffer.Clear();
                }
            }

            for (var current = inline; current != null; current = current.NextSibling) {
                if (current is LinkInline link) {
                    Flush();
                    if (link.IsImage) {
                        AddImage(document, paragraph, link);
                    } else {
                        string label = BuildMarkdown(link.FirstChild);
                        var hyperlink = paragraph.AddHyperLink(label, new Uri(link.Url, UriKind.RelativeOrAbsolute));
                        if (!string.IsNullOrEmpty(options.FontFamily)) {
                            hyperlink.SetFontFamily(options.FontFamily);
                        }
                    }
                } else {
                    buffer.Append(BuildMarkdown(current));
                }
            }
            Flush();
        }

        private static void AddImage(WordDocument document, WordParagraph paragraph, LinkInline link) {
            if (File.Exists(link.Url)) {
                paragraph.AddImage(link.Url);
            } else {
                document.AddImageFromUrl(link.Url, 50, 50);
            }
        }

        private static string BuildMarkdown(Inline? inline) {
            if (inline == null) {
                return string.Empty;
            }

            var sb = new StringBuilder();
            for (var current = inline; current != null; current = current.NextSibling) {
                switch (current) {
                    case LiteralInline literal:
                        sb.Append(literal.Content.ToString());
                        break;
                    case EmphasisInline emphasis:
                        string marker = new('*', emphasis.DelimiterCount);
                        sb.Append(marker);
                        sb.Append(BuildMarkdown(emphasis.FirstChild));
                        sb.Append(marker);
                        break;
                    case LineBreakInline:
                        sb.Append('\n');
                        break;
                    case ContainerInline container:
                        sb.Append(BuildMarkdown(container.FirstChild));
                        break;
                }
            }

            return sb.ToString();
        }

        private static string GetCodeBlockText(CodeBlock codeBlock) {
            var sb = new StringBuilder();
            foreach (var line in codeBlock.Lines.Lines) {
                sb.AppendLine(line.Slice.ToString());
            }
            return sb.ToString().TrimEnd();
        }

    }
}