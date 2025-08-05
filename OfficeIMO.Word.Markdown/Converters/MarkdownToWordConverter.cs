using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

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
            if (markdown == null) throw new ArgumentNullException(nameof(markdown));
            options ??= new MarkdownToWordOptions();
            
            var wordDoc = WordDocument.Create();
            
            // Apply defaults from options
            if (options.DefaultPageSize.HasValue) {
                wordDoc.PageSettings.PageSize = options.DefaultPageSize.Value;
            }
            if (options.DefaultOrientation.HasValue) {
                wordDoc.PageOrientation = options.DefaultOrientation.Value;
            }
            
            // TODO: Implement full Markdown to Word conversion using Markdig
            // For now, just add the markdown as plain text
            
            var lines = markdown.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (var line in lines) {
                if (!string.IsNullOrWhiteSpace(line)) {
                    wordDoc.AddParagraph(line);
                } else {
                    wordDoc.AddParagraph();
                }
            }
            
            return wordDoc;
        }
    }
}