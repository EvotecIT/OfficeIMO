using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html.Converters {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Use OfficeIMO.Word API properties/methods to read document content
    /// 2. Use document.Paragraphs, document.Tables, document.Lists, etc.
    /// 3. For formatting, use paragraph.Bold, paragraph.Italic, etc.
    /// 4. For styles, use paragraph.Style (WordParagraphStyles enum)
    /// 5. Only access OpenXML internals if OfficeIMO.Word API doesn't expose needed data
    /// </summary>
    internal class WordToHtmlConverter {
        public async Task<string> ConvertAsync(WordDocument document, WordToHtmlOptions options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToHtmlOptions();
            
            var config = Configuration.Default;
            var context = BrowsingContext.New(config);
            var htmlDoc = await context.OpenNewAsync();
            
            // Set up HTML document
            var head = htmlDoc.Head;
            var body = htmlDoc.Body;
            
            // Add meta tags
            var charset = htmlDoc.CreateElement("meta");
            charset.SetAttribute("charset", "UTF-8");
            head.AppendChild(charset);
            
            // Add title
            var title = htmlDoc.CreateElement("title");
            title.TextContent = "Document"; // TODO: Get title from document properties
            head.AppendChild(title);
            
            // TODO: Implement full Word to HTML conversion
            // For now, just add paragraphs as simple <p> elements
            
            foreach (var paragraph in document.Paragraphs) {
                if (!string.IsNullOrEmpty(paragraph.Text)) {
                    var p = htmlDoc.CreateElement("p");
                    p.TextContent = paragraph.Text;
                    body.AppendChild(p);
                }
            }
            
            return htmlDoc.DocumentElement.OuterHtml;
        }
    }
}