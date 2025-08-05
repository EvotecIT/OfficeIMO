using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.Linq;

namespace OfficeIMO.Word.Markdown.Converters {
    internal class MarkdownToWordConverter {
        public WordDocument Convert(string markdown, MarkdownOptions options) {
            options ??= new MarkdownOptions();
            
            var pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .UsePipeTables()
                .UseGridTables()
                .UseEmphasisExtras()
                .UseAutoLinks()
                .UseTaskLists()
                .Build();
            
            var markdownDoc = Markdig.Markdown.Parse(markdown, pipeline);
            var wordDoc = WordDocument.Create();
            
            // Walk through the markdown document blocks
            foreach (var block in markdownDoc) {
                ProcessBlock(wordDoc, block, options);
            }
            
            return wordDoc;
        }
        
        private void ProcessBlock(WordDocument doc, Block block, MarkdownOptions options) {
            switch (block) {
                case HeadingBlock heading:
                    ProcessHeading(doc, heading, options);
                    break;
                    
                case ParagraphBlock paragraph:
                    ProcessParagraph(doc, paragraph, options);
                    break;
                    
                case ListBlock list:
                    ProcessList(doc, list, options);
                    break;
                    
                case CodeBlock code:
                    ProcessCodeBlock(doc, code, options);
                    break;
                    
                case Table table:
                    ProcessTable(doc, table, options);
                    break;
                    
                case QuoteBlock quote:
                    ProcessQuote(doc, quote, options);
                    break;
                    
                case ThematicBreakBlock:
                    doc.AddHorizontalLine();
                    break;
            }
        }
        
        private void ProcessHeading(WordDocument doc, HeadingBlock heading, MarkdownOptions options) {
            var paragraph = doc.AddParagraph();
            ProcessInlines(paragraph, heading.Inline, options);
            
            paragraph.Style = heading.Level switch {
                1 => WordParagraphStyles.Heading1,
                2 => WordParagraphStyles.Heading2,
                3 => WordParagraphStyles.Heading3,
                4 => WordParagraphStyles.Heading4,
                5 => WordParagraphStyles.Heading5,
                6 => WordParagraphStyles.Heading6,
                _ => WordParagraphStyles.Normal
            };
        }
        
        private void ProcessParagraph(WordDocument doc, ParagraphBlock paragraph, MarkdownOptions options) {
            var wordParagraph = doc.AddParagraph();
            ProcessInlines(wordParagraph, paragraph.Inline, options);
        }
        
        private void ProcessInlines(WordParagraph paragraph, ContainerInline inlines, MarkdownOptions options) {
            if (inlines == null) return;
            
            foreach (var inline in inlines) {
                ProcessInline(paragraph, inline, options);
            }
        }
        
        private void ProcessInline(WordParagraph paragraph, Inline inline, MarkdownOptions options) {
            switch (inline) {
                case LiteralInline literal:
                    paragraph.AddText(literal.Content.ToString());
                    break;
                    
                case EmphasisInline emphasis:
                    var text = GetInlineText(emphasis);
                    if (emphasis.DelimiterCount == 2) {
                        paragraph.AddText(text).Bold = true;
                    } else {
                        paragraph.AddText(text).Italic = true;
                    }
                    break;
                    
                case LinkInline link:
                    if (link.IsImage) {
                        // TODO: Download and add image
                        paragraph.AddText($"[Image: {link.Title ?? link.Url}]");
                    } else {
                        paragraph.AddHyperlink(GetInlineText(link), new Uri(link.Url));
                    }
                    break;
                    
                case CodeInline code:
                    var codeText = paragraph.AddText(code.Content);
                    codeText.FontFamily = "Consolas";
                    codeText.Highlight = HighlightColor.LightGray;
                    break;
                    
                case LineBreakInline:
                    paragraph.AddText("\n");
                    break;
            }
        }
        
        private string GetInlineText(ContainerInline container) {
            var text = "";
            foreach (var child in container) {
                if (child is LiteralInline literal) {
                    text += literal.Content.ToString();
                }
            }
            return text;
        }
        
        private void ProcessList(WordDocument doc, ListBlock list, MarkdownOptions options) {
            var wordList = doc.AddList(list.IsOrdered ? WordListStyle.Heading1ai : WordListStyle.Bulleted);
            
            foreach (var item in list) {
                if (item is ListItemBlock listItem) {
                    var firstBlock = listItem.FirstOrDefault();
                    if (firstBlock is ParagraphBlock para) {
                        var text = GetInlineText(para.Inline);
                        wordList.AddItem(text);
                    }
                }
            }
        }
        
        private void ProcessCodeBlock(WordDocument doc, CodeBlock code, MarkdownOptions options) {
            var paragraph = doc.AddParagraph();
            paragraph.AddText(code.Lines.ToString());
            paragraph.FontFamily = "Consolas";
            paragraph.Highlight = HighlightColor.LightGray;
        }
        
        private void ProcessTable(WordDocument doc, Table table, MarkdownOptions options) {
            // Count columns
            var columnCount = 0;
            foreach (var row in table) {
                if (row is TableRow tableRow) {
                    columnCount = Math.Max(columnCount, tableRow.Count);
                }
            }
            
            if (columnCount == 0) return;
            
            var rowCount = table.Count;
            var wordTable = doc.AddTable(rowCount, columnCount);
            
            int rowIndex = 0;
            foreach (var row in table) {
                if (row is TableRow tableRow) {
                    int colIndex = 0;
                    foreach (var cell in tableRow) {
                        if (cell is TableCell tableCell && colIndex < columnCount) {
                            var cellPara = wordTable.Rows[rowIndex].Cells[colIndex].Paragraphs[0];
                            ProcessInlines(cellPara, tableCell.Inline, options);
                            colIndex++;
                        }
                    }
                    rowIndex++;
                }
            }
        }
        
        private void ProcessQuote(WordDocument doc, QuoteBlock quote, MarkdownOptions options) {
            var paragraph = doc.AddParagraph();
            paragraph.Indentation.Left = 720; // 0.5 inch
            paragraph.Italic = true;
            
            foreach (var block in quote) {
                if (block is ParagraphBlock para) {
                    ProcessInlines(paragraph, para.Inline, options);
                }
            }
        }
    }
}