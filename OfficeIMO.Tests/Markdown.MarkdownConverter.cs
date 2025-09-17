using OfficeIMO.Word.Markdown;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact(Skip = "TODO: Implement additional Markdown to Word scenarios with OfficeIMO.Markdown reader")]
        public void Test_Markdown_RoundTrip() {
            string md = "# Heading 1\n\nHello **world** and *universe*.";

            var doc = md.LoadFromMarkdown( new MarkdownToWordOptions { FontFamily = "Calibri" });
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("# Heading 1", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("**world**", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("*universe*", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact(Skip = "TODO: Implement Markdown list parsing and conversion to WordList")]
        public void Test_Markdown_Lists_RoundTrip() {
            string md = "- Item 1\n- Item 2\n\n1. First\n1. Second";

            var doc = md.LoadFromMarkdown( new MarkdownToWordOptions { FontFamily = "Calibri" });
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("- Item 1", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("1. First", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact(Skip = "TODO: Implement font family resolution for Markdown")]
        public void Test_Markdown_FontResolver() {
            string md = "Hello";
            using MemoryStream ms = new MemoryStream();

            var doc = md.LoadFromMarkdown( new MarkdownToWordOptions { FontFamily = "monospace" });
            doc.Save(ms);

            ms.Position = 0;
            using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
            RunFonts fonts = docx.MainDocumentPart!.Document.Body!.Descendants<RunFonts>().First();
            Assert.Equal(FontResolver.Resolve("monospace"), fonts.Ascii);
        }

        [Fact(Skip = "TODO: Implement automatic URL detection in Markdown text")]
        public void Test_Markdown_Urls_CreateHyperlinks() {
            string md = "Visit http://example.com";
            using MemoryStream ms = new MemoryStream();

            var doc = md.LoadFromMarkdown( new MarkdownToWordOptions());
            doc.Save(ms);

            ms.Position = 0;
            using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
            var hyperlink = docx.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>().FirstOrDefault();
            Assert.NotNull(hyperlink);
            var rel = docx.MainDocumentPart.HyperlinkRelationships.First();
            Assert.StartsWith("http://example.com", rel.Uri.ToString());
        }

        [Fact]
        public void Test_Markdown_HtmlBlock_RoundTrip() {
            string md = "<p><strong>Bold</strong> HTML</p>";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("**Bold** HTML", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToMarkdown_PreservesEmbeddedHtml() {
            using var doc = WordDocument.Create();
            doc.AddEmbeddedFragment("<div>HTML Block</div>", WordAlternativeFormatImportPartType.Html);
            string? html = doc.EmbeddedDocuments[0].GetHtml();
            Assert.NotNull(html);
            Assert.Contains("<div>HTML Block</div>", html, StringComparison.OrdinalIgnoreCase);
            string md = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("<div>HTML Block</div>", md, StringComparison.OrdinalIgnoreCase);
        }
    }
}
