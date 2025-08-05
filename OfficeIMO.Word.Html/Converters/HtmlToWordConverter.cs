using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using OfficeIMO.Word;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html.Converters {
    internal class HtmlToWordConverter {
        public async Task<WordDocument> ConvertAsync(string html, HtmlOptions options) {
            options ??= new HtmlOptions();
            
            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            var parser = context.GetService<IHtmlParser>();
            var document = await parser.ParseDocumentAsync(html);
            
            var wordDoc = WordDocument.Create();
            ProcessElement(wordDoc, document.Body, options);
            
            return wordDoc;
        }
        
        public WordDocument Convert(string html, HtmlOptions options) {
            return ConvertAsync(html, options).GetAwaiter().GetResult();
        }
        
        private void ProcessElement(WordDocument doc, IElement element, HtmlOptions options) {
            if (element == null) return;
            
            foreach (var child in element.Children) {
                ProcessNode(doc, child, options, null);
            }
        }
        
        private void ProcessNode(WordDocument doc, INode node, HtmlOptions options, WordParagraph currentParagraph) {
            switch (node) {
                case IHtmlHeadingElement heading:
                    ProcessHeading(doc, heading, options);
                    break;
                    
                case IHtmlParagraphElement paragraph:
                    ProcessParagraph(doc, paragraph, options);
                    break;
                    
                case IHtmlDivElement div:
                    ProcessDiv(doc, div, options);
                    break;
                    
                case IHtmlUnorderedListElement ul:
                    ProcessUnorderedList(doc, ul, options);
                    break;
                    
                case IHtmlOrderedListElement ol:
                    ProcessOrderedList(doc, ol, options);
                    break;
                    
                case IHtmlTableElement table:
                    ProcessTable(doc, table, options);
                    break;
                    
                case IHtmlPreElement pre:
                    ProcessPreformatted(doc, pre, options);
                    break;
                    
                case IHtmlQuoteElement quote:
                    ProcessBlockquote(doc, quote, options);
                    break;
                    
                case IHtmlHrElement:
                    doc.AddHorizontalLine();
                    break;
                    
                case IHtmlImageElement img:
                    ProcessImage(doc, img, options, currentParagraph);
                    break;
                    
                case IHtmlAnchorElement anchor:
                    ProcessAnchor(doc, anchor, options, currentParagraph);
                    break;
                    
                case IElement element:
                    // Process children of unknown elements
                    foreach (var child in element.ChildNodes) {
                        ProcessNode(doc, child, options, currentParagraph);
                    }
                    break;
                    
                case IText text:
                    if (currentParagraph != null && !string.IsNullOrWhiteSpace(text.TextContent)) {
                        currentParagraph.AddText(text.TextContent);
                    }
                    break;
            }
        }
        
        private void ProcessHeading(WordDocument doc, IHtmlHeadingElement heading, HtmlOptions options) {
            var paragraph = doc.AddParagraph();
            ProcessInlineContent(paragraph, heading, options);
            
            var level = heading.LocalName.Substring(1); // h1 -> 1, h2 -> 2, etc.
            paragraph.Style = level switch {
                "1" => WordParagraphStyles.Heading1,
                "2" => WordParagraphStyles.Heading2,
                "3" => WordParagraphStyles.Heading3,
                "4" => WordParagraphStyles.Heading4,
                "5" => WordParagraphStyles.Heading5,
                "6" => WordParagraphStyles.Heading6,
                _ => WordParagraphStyles.Normal
            };
        }
        
        private void ProcessParagraph(WordDocument doc, IHtmlParagraphElement paragraph, HtmlOptions options) {
            var wordParagraph = doc.AddParagraph();
            ProcessInlineContent(wordParagraph, paragraph, options);
            
            // Apply styles if present
            if (options.PreserveStyles) {
                var style = paragraph.GetStyle();
                ApplyStyles(wordParagraph, style);
            }
        }
        
        private void ProcessDiv(WordDocument doc, IHtmlDivElement div, HtmlOptions options) {
            // Treat div as a paragraph container
            var paragraph = doc.AddParagraph();
            ProcessInlineContent(paragraph, div, options);
        }
        
        private void ProcessInlineContent(WordParagraph paragraph, IElement element, HtmlOptions options) {
            foreach (var node in element.ChildNodes) {
                ProcessInlineNode(paragraph, node, options);
            }
        }
        
        private void ProcessInlineNode(WordParagraph paragraph, INode node, HtmlOptions options) {
            switch (node) {
                case IHtmlElement htmlElement when htmlElement.LocalName == "b" || htmlElement.LocalName == "strong":
                    var boldText = paragraph.AddText(htmlElement.TextContent);
                    boldText.Bold = true;
                    break;
                    
                case IHtmlElement htmlElement when htmlElement.LocalName == "i" || htmlElement.LocalName == "em":
                    var italicText = paragraph.AddText(htmlElement.TextContent);
                    italicText.Italic = true;
                    break;
                    
                case IHtmlElement htmlElement when htmlElement.LocalName == "u":
                    var underlineText = paragraph.AddText(htmlElement.TextContent);
                    underlineText.Underline = UnderlineValues.Single;
                    break;
                    
                case IHtmlElement htmlElement when htmlElement.LocalName == "code":
                    var codeText = paragraph.AddText(htmlElement.TextContent);
                    codeText.FontFamily = options.CodeFontFamily;
                    codeText.Highlight = HighlightColor.LightGray;
                    break;
                    
                case IHtmlAnchorElement anchor:
                    if (!string.IsNullOrEmpty(anchor.Href)) {
                        paragraph.AddHyperlink(anchor.TextContent, new Uri(anchor.Href, UriKind.RelativeOrAbsolute));
                    } else {
                        paragraph.AddText(anchor.TextContent);
                    }
                    break;
                    
                case IHtmlBrElement:
                    paragraph.AddText("\n");
                    break;
                    
                case IHtmlSpanElement span:
                    var spanText = paragraph.AddText(span.TextContent);
                    if (options.PreserveStyles) {
                        var style = span.GetStyle();
                        ApplyInlineStyles(spanText, style);
                    }
                    break;
                    
                case IText text:
                    if (!string.IsNullOrWhiteSpace(text.TextContent)) {
                        paragraph.AddText(text.TextContent);
                    }
                    break;
                    
                case IElement element:
                    // Recursively process children
                    foreach (var child in element.ChildNodes) {
                        ProcessInlineNode(paragraph, child, options);
                    }
                    break;
            }
        }
        
        private void ProcessUnorderedList(WordDocument doc, IHtmlUnorderedListElement ul, HtmlOptions options) {
            var list = doc.AddList(WordListStyle.Bulleted);
            
            foreach (var item in ul.Children.OfType<IHtmlListItemElement>()) {
                var listItem = list.AddItem(item.TextContent.Trim());
                // TODO: Handle nested lists
            }
        }
        
        private void ProcessOrderedList(WordDocument doc, IHtmlOrderedListElement ol, HtmlOptions options) {
            var list = doc.AddList(WordListStyle.Heading1ai);
            
            foreach (var item in ol.Children.OfType<IHtmlListItemElement>()) {
                var listItem = list.AddItem(item.TextContent.Trim());
                // TODO: Handle nested lists
            }
        }
        
        private void ProcessTable(WordDocument doc, IHtmlTableElement table, HtmlOptions options) {
            var rows = table.Rows.Count();
            var cols = table.Rows.FirstOrDefault()?.Cells.Count() ?? 0;
            
            if (rows == 0 || cols == 0) return;
            
            var wordTable = doc.AddTable(rows, cols);
            
            int rowIndex = 0;
            foreach (var row in table.Rows) {
                int colIndex = 0;
                foreach (var cell in row.Cells) {
                    if (colIndex < cols && rowIndex < rows) {
                        var wordCell = wordTable.Rows[rowIndex].Cells[colIndex];
                        wordCell.Paragraphs[0].Text = cell.TextContent.Trim();
                        
                        // Handle colspan and rowspan
                        var colspan = cell.GetAttribute("colspan");
                        var rowspan = cell.GetAttribute("rowspan");
                        // TODO: Implement cell merging
                        
                        colIndex++;
                    }
                }
                rowIndex++;
            }
        }
        
        private void ProcessPreformatted(WordDocument doc, IHtmlPreElement pre, HtmlOptions options) {
            var paragraph = doc.AddParagraph();
            paragraph.AddText(pre.TextContent);
            paragraph.FontFamily = options.CodeFontFamily;
            paragraph.Highlight = HighlightColor.LightGray;
        }
        
        private void ProcessBlockquote(WordDocument doc, IHtmlQuoteElement quote, HtmlOptions options) {
            var paragraph = doc.AddParagraph();
            paragraph.Indentation.Left = 720; // 0.5 inch
            paragraph.Italic = true;
            ProcessInlineContent(paragraph, quote, options);
        }
        
        private void ProcessImage(WordDocument doc, IHtmlImageElement img, HtmlOptions options, WordParagraph paragraph) {
            if (paragraph == null) {
                paragraph = doc.AddParagraph();
            }
            
            var src = img.Source;
            var alt = img.AlternativeText ?? "Image";
            
            if (options.DownloadImages && !string.IsNullOrEmpty(src)) {
                // TODO: Download image and add to document
                paragraph.AddText($"[Image: {alt}]");
            } else {
                paragraph.AddText($"[Image: {alt}]");
            }
        }
        
        private void ProcessAnchor(WordDocument doc, IHtmlAnchorElement anchor, HtmlOptions options, WordParagraph paragraph) {
            if (paragraph == null) {
                paragraph = doc.AddParagraph();
            }
            
            if (!string.IsNullOrEmpty(anchor.Href)) {
                paragraph.AddHyperlink(anchor.TextContent, new Uri(anchor.Href, UriKind.RelativeOrAbsolute));
            } else {
                paragraph.AddText(anchor.TextContent);
            }
        }
        
        private void ApplyStyles(WordParagraph paragraph, ICssStyleDeclaration style) {
            // TODO: Parse and apply CSS styles
            // text-align, font-size, color, etc.
        }
        
        private void ApplyInlineStyles(WordText text, ICssStyleDeclaration style) {
            // TODO: Parse and apply inline CSS styles
            // font-weight, font-style, color, etc.
        }
    }
}