using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html.Converters {
    internal class WordToHtmlConverter {
        public async Task<string> ConvertAsync(WordDocument document, HtmlOptions options) {
            options ??= new HtmlOptions();
            
            var config = Configuration.Default;
            var context = BrowsingContext.New(config);
            var htmlDoc = await context.OpenNewAsync();
            
            // Set up HTML document
            var html = htmlDoc.DocumentElement as IHtmlHtmlElement;
            var head = htmlDoc.Head;
            var body = htmlDoc.Body;
            
            // Add meta tags
            var charset = htmlDoc.CreateElement("meta");
            charset.SetAttribute("charset", "UTF-8");
            head.AppendChild(charset);
            
            var viewport = htmlDoc.CreateElement("meta");
            viewport.SetAttribute("name", "viewport");
            viewport.SetAttribute("content", "width=device-width, initial-scale=1.0");
            head.AppendChild(viewport);
            
            // Add title
            var title = htmlDoc.CreateElement("title");
            title.TextContent = document.Title ?? "Document";
            head.AppendChild(title);
            
            // Add CSS if requested
            if (options.IncludeCss) {
                var style = htmlDoc.CreateElement("style");
                style.TextContent = GetDefaultCss(options);
                head.AppendChild(style);
            }
            
            // Process document elements
            foreach (var element in document.Elements) {
                ProcessElement(htmlDoc, body, element, options);
            }
            
            return htmlDoc.DocumentElement.OuterHtml;
        }
        
        public string Convert(WordDocument document, HtmlOptions options) {
            return ConvertAsync(document, options).GetAwaiter().GetResult();
        }
        
        private void ProcessElement(IHtmlDocument htmlDoc, IElement parent, WordElement element, HtmlOptions options) {
            switch (element) {
                case WordParagraph paragraph:
                    ProcessParagraph(htmlDoc, parent, paragraph, options);
                    break;
                    
                case WordTable table:
                    ProcessTable(htmlDoc, parent, table, options);
                    break;
                    
                case WordList list:
                    ProcessList(htmlDoc, parent, list, options);
                    break;
                    
                case WordPageBreak:
                    var hr = htmlDoc.CreateElement("hr");
                    parent.AppendChild(hr);
                    break;
            }
        }
        
        private void ProcessParagraph(IHtmlDocument htmlDoc, IElement parent, WordParagraph paragraph, HtmlOptions options) {
            IElement element;
            
            // Create appropriate element based on style
            if (paragraph.Style != null) {
                switch (paragraph.Style) {
                    case WordParagraphStyles.Heading1:
                        element = htmlDoc.CreateElement("h1");
                        break;
                    case WordParagraphStyles.Heading2:
                        element = htmlDoc.CreateElement("h2");
                        break;
                    case WordParagraphStyles.Heading3:
                        element = htmlDoc.CreateElement("h3");
                        break;
                    case WordParagraphStyles.Heading4:
                        element = htmlDoc.CreateElement("h4");
                        break;
                    case WordParagraphStyles.Heading5:
                        element = htmlDoc.CreateElement("h5");
                        break;
                    case WordParagraphStyles.Heading6:
                        element = htmlDoc.CreateElement("h6");
                        break;
                    default:
                        element = htmlDoc.CreateElement("p");
                        break;
                }
            } else {
                element = htmlDoc.CreateElement("p");
            }
            
            // Apply paragraph styles
            if (options.PreserveStyles) {
                var styleStr = "";
                
                if (paragraph.Alignment == JustificationValues.Center) {
                    styleStr += "text-align: center;";
                } else if (paragraph.Alignment == JustificationValues.Right) {
                    styleStr += "text-align: right;";
                } else if (paragraph.Alignment == JustificationValues.Both) {
                    styleStr += "text-align: justify;";
                }
                
                if (paragraph.Indentation?.Left > 0) {
                    var indent = paragraph.Indentation.Left / 1440.0; // Convert twips to inches
                    styleStr += $"margin-left: {indent}in;";
                }
                
                if (!string.IsNullOrEmpty(styleStr)) {
                    element.SetAttribute("style", styleStr.TrimEnd(';'));
                }
            }
            
            // Process runs
            foreach (var run in paragraph.Runs) {
                ProcessRun(htmlDoc, element, run, options);
            }
            
            // Process hyperlinks
            foreach (var hyperlink in paragraph.Hyperlinks) {
                var anchor = htmlDoc.CreateElement("a") as IHtmlAnchorElement;
                anchor.Href = hyperlink.Uri?.ToString() ?? "";
                anchor.TextContent = hyperlink.Text ?? hyperlink.Uri?.ToString() ?? "";
                element.AppendChild(anchor);
            }
            
            // Process images
            foreach (var image in paragraph.Images) {
                var img = htmlDoc.CreateElement("img") as IHtmlImageElement;
                // TODO: Handle image export
                img.AlternativeText = image.Description ?? "Image";
                img.SetAttribute("src", "data:image/png;base64,"); // Placeholder
                element.AppendChild(img);
            }
            
            parent.AppendChild(element);
        }
        
        private void ProcessRun(IHtmlDocument htmlDoc, IElement parent, WordRun run, HtmlOptions options) {
            var text = run.Text;
            if (string.IsNullOrEmpty(text)) return;
            
            // Apply formatting
            IElement element = null;
            
            if (run.Bold && run.Italic) {
                var strong = htmlDoc.CreateElement("strong");
                var em = htmlDoc.CreateElement("em");
                em.TextContent = text;
                strong.AppendChild(em);
                element = strong;
            } else if (run.Bold) {
                element = htmlDoc.CreateElement("strong");
                element.TextContent = text;
            } else if (run.Italic) {
                element = htmlDoc.CreateElement("em");
                element.TextContent = text;
            } else if (run.Underline != null && run.Underline != UnderlineValues.None) {
                element = htmlDoc.CreateElement("u");
                element.TextContent = text;
            } else if (run.IsCode) {
                element = htmlDoc.CreateElement("code");
                element.TextContent = text;
            } else {
                // Plain text or span with styles
                if (options.PreserveStyles && HasSpecialFormatting(run)) {
                    element = htmlDoc.CreateElement("span");
                    element.TextContent = text;
                    ApplyRunStyles(element, run, options);
                } else {
                    parent.AppendChild(htmlDoc.CreateTextNode(text));
                    return;
                }
            }
            
            if (element != null) {
                parent.AppendChild(element);
            }
        }
        
        private void ProcessList(IHtmlDocument htmlDoc, IElement parent, WordList list, HtmlOptions options) {
            var listElement = list.ListType == WordListStyle.Bulleted
                ? htmlDoc.CreateElement("ul")
                : htmlDoc.CreateElement("ol");
            
            foreach (var item in list.Items) {
                var li = htmlDoc.CreateElement("li");
                
                // Process runs in the list item
                foreach (var run in item.Runs) {
                    ProcessRun(htmlDoc, li, run, options);
                }
                
                listElement.AppendChild(li);
            }
            
            parent.AppendChild(listElement);
        }
        
        private void ProcessTable(IHtmlDocument htmlDoc, IElement parent, WordTable table, HtmlOptions options) {
            var tableElement = htmlDoc.CreateElement("table");
            
            if (options.IncludeCss) {
                tableElement.SetAttribute("style", "border-collapse: collapse; width: 100%;");
            }
            
            for (int rowIndex = 0; rowIndex < table.RowsCount; rowIndex++) {
                var tr = htmlDoc.CreateElement("tr");
                var row = table.Rows[rowIndex];
                
                for (int colIndex = 0; colIndex < row.CellsCount; colIndex++) {
                    var cell = row.Cells[colIndex];
                    var td = rowIndex == 0 ? htmlDoc.CreateElement("th") : htmlDoc.CreateElement("td");
                    
                    if (options.IncludeCss) {
                        td.SetAttribute("style", "border: 1px solid #ddd; padding: 8px;");
                    }
                    
                    // Process cell paragraphs
                    foreach (var paragraph in cell.Paragraphs) {
                        ProcessParagraph(htmlDoc, td, paragraph, options);
                    }
                    
                    // Handle cell merging
                    if (cell.GridSpan > 1) {
                        td.SetAttribute("colspan", cell.GridSpan.ToString());
                    }
                    
                    tr.AppendChild(td);
                }
                
                tableElement.AppendChild(tr);
            }
            
            parent.AppendChild(tableElement);
        }
        
        private bool HasSpecialFormatting(WordRun run) {
            return run.Color != null || 
                   run.FontSize != null || 
                   run.FontFamily != null ||
                   run.Highlight != null;
        }
        
        private void ApplyRunStyles(IElement element, WordRun run, HtmlOptions options) {
            var styleStr = "";
            
            if (run.Color != null) {
                styleStr += $"color: {run.Color};";
            }
            
            if (run.FontSize != null) {
                styleStr += $"font-size: {run.FontSize}pt;";
            }
            
            if (run.FontFamily != null) {
                styleStr += $"font-family: '{run.FontFamily}';";
            }
            
            if (run.Highlight != null) {
                styleStr += $"background-color: {GetHighlightColor(run.Highlight.Value)};";
            }
            
            if (!string.IsNullOrEmpty(styleStr)) {
                element.SetAttribute("style", styleStr.TrimEnd(';'));
            }
        }
        
        private string GetHighlightColor(HighlightColorValues highlight) {
            return highlight switch {
                HighlightColorValues.Yellow => "yellow",
                HighlightColorValues.Green => "lightgreen",
                HighlightColorValues.Cyan => "cyan",
                HighlightColorValues.Magenta => "magenta",
                HighlightColorValues.Blue => "lightblue",
                HighlightColorValues.Red => "salmon",
                HighlightColorValues.DarkBlue => "darkblue",
                HighlightColorValues.DarkCyan => "darkcyan",
                HighlightColorValues.DarkGreen => "darkgreen",
                HighlightColorValues.DarkMagenta => "darkmagenta",
                HighlightColorValues.DarkRed => "darkred",
                HighlightColorValues.DarkYellow => "gold",
                HighlightColorValues.DarkGray => "darkgray",
                HighlightColorValues.LightGray => "lightgray",
                HighlightColorValues.Black => "black",
                _ => "yellow"
            };
        }
        
        private string GetDefaultCss(HtmlOptions options) {
            return @"
body {
    font-family: '" + options.FontFamily + @"', sans-serif;
    line-height: 1.6;
    margin: 20px;
    max-width: 800px;
}

h1, h2, h3, h4, h5, h6 {
    margin-top: 1em;
    margin-bottom: 0.5em;
}

p {
    margin-bottom: 1em;
}

table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 1em;
}

th, td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}

th {
    background-color: #f2f2f2;
    font-weight: bold;
}

code {
    font-family: '" + options.CodeFontFamily + @"', monospace;
    background-color: #f4f4f4;
    padding: 2px 4px;
    border-radius: 3px;
}

pre {
    font-family: '" + options.CodeFontFamily + @"', monospace;
    background-color: #f4f4f4;
    padding: 10px;
    border-radius: 5px;
    overflow-x: auto;
}

blockquote {
    margin: 1em 0;
    padding-left: 1em;
    border-left: 3px solid #ddd;
    font-style: italic;
}

a {
    color: #0066cc;
    text-decoration: none;
}

a:hover {
    text-decoration: underline;
}
";
        }
    }
}