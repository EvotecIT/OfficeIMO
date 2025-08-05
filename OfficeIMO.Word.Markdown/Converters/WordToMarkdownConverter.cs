using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown.Converters {
    internal class WordToMarkdownConverter {
        private readonly StringBuilder _output = new StringBuilder();
        private readonly HashSet<string> _processedListIds = new HashSet<string>();
        
        public string Convert(WordDocument document, MarkdownOptions options) {
            options ??= new MarkdownOptions();
            
            foreach (var element in document.Elements) {
                ProcessElement(element, options);
            }
            
            return _output.ToString().TrimEnd();
        }
        
        private void ProcessElement(WordElement element, MarkdownOptions options) {
            switch (element) {
                case WordParagraph paragraph:
                    ProcessParagraph(paragraph, options);
                    break;
                    
                case WordTable table:
                    ProcessTable(table, options);
                    break;
                    
                case WordList list:
                    ProcessList(list, options);
                    break;
                    
                case WordPageBreak:
                    _output.AppendLine("\n---\n");
                    break;
            }
        }
        
        private void ProcessParagraph(WordParagraph paragraph, MarkdownOptions options) {
            // Skip if this paragraph is part of a list we've already processed
            if (paragraph.IsListItem && _processedListIds.Contains(paragraph.ListId)) {
                return;
            }
            
            var text = GetParagraphMarkdown(paragraph);
            
            // Handle headings
            if (paragraph.Style != null) {
                switch (paragraph.Style) {
                    case WordParagraphStyles.Heading1:
                        _output.AppendLine($"# {text}");
                        break;
                    case WordParagraphStyles.Heading2:
                        _output.AppendLine($"## {text}");
                        break;
                    case WordParagraphStyles.Heading3:
                        _output.AppendLine($"### {text}");
                        break;
                    case WordParagraphStyles.Heading4:
                        _output.AppendLine($"#### {text}");
                        break;
                    case WordParagraphStyles.Heading5:
                        _output.AppendLine($"##### {text}");
                        break;
                    case WordParagraphStyles.Heading6:
                        _output.AppendLine($"###### {text}");
                        break;
                    default:
                        _output.AppendLine(text);
                        break;
                }
            } else {
                _output.AppendLine(text);
            }
            
            if (options.PreserveEmptyLines && !string.IsNullOrWhiteSpace(text)) {
                _output.AppendLine();
            }
        }
        
        private string GetParagraphMarkdown(WordParagraph paragraph) {
            var sb = new StringBuilder();
            
            // Process runs
            foreach (var run in paragraph.Runs) {
                var text = run.Text;
                
                // Apply formatting
                if (run.Bold && run.Italic) {
                    text = $"***{text}***";
                } else if (run.Bold) {
                    text = $"**{text}**";
                } else if (run.Italic) {
                    text = $"*{text}*";
                }
                
                if (run.IsCode) {
                    text = $"`{run.Text}`";
                }
                
                sb.Append(text);
            }
            
            // Process hyperlinks
            foreach (var hyperlink in paragraph.Hyperlinks) {
                var linkText = hyperlink.Text ?? hyperlink.Uri?.ToString() ?? "";
                var linkUrl = hyperlink.Uri?.ToString() ?? "";
                sb.Append($"[{linkText}]({linkUrl})");
            }
            
            // Process images
            foreach (var image in paragraph.Images) {
                var altText = image.Description ?? "Image";
                // TODO: Handle image export/path
                sb.Append($"![{altText}](image-placeholder)");
            }
            
            return sb.ToString();
        }
        
        private void ProcessList(WordList list, MarkdownOptions options) {
            _processedListIds.Add(list.ListId);
            
            int index = 1;
            foreach (var item in list.Items) {
                var prefix = list.ListType == WordListStyle.Bulleted ? "- " : $"{index}. ";
                var text = GetParagraphMarkdown(item);
                _output.AppendLine($"{prefix}{text}");
                index++;
            }
            
            _output.AppendLine();
        }
        
        private void ProcessTable(WordTable table, MarkdownOptions options) {
            if (table.RowsCount == 0 || table.ColumnsCount == 0) return;
            
            // Process header row (first row)
            _output.Append("|");
            for (int col = 0; col < table.ColumnsCount; col++) {
                var cell = table.Rows[0].Cells[col];
                var text = GetCellText(cell);
                _output.Append($" {text} |");
            }
            _output.AppendLine();
            
            // Add separator row
            _output.Append("|");
            for (int col = 0; col < table.ColumnsCount; col++) {
                _output.Append(" --- |");
            }
            _output.AppendLine();
            
            // Process data rows
            for (int row = 1; row < table.RowsCount; row++) {
                _output.Append("|");
                for (int col = 0; col < table.ColumnsCount; col++) {
                    var cell = table.Rows[row].Cells[col];
                    var text = GetCellText(cell);
                    _output.Append($" {text} |");
                }
                _output.AppendLine();
            }
            
            _output.AppendLine();
        }
        
        private string GetCellText(WordTableCell cell) {
            var sb = new StringBuilder();
            foreach (var paragraph in cell.Paragraphs) {
                if (sb.Length > 0) sb.Append(" ");
                sb.Append(GetParagraphMarkdown(paragraph));
            }
            return sb.ToString();
        }
    }
}