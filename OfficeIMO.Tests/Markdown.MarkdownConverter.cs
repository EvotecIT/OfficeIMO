using OfficeIMO.Markdown;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Test_Markdown_RoundTrip() {
            string md = "# Heading 1\n\nHello **world** and *universe*.";
            using MemoryStream ms = new MemoryStream();
            MarkdownToWordConverter.Convert(md, ms, new MarkdownToWordOptions { FontFamily = "Calibri" });

            ms.Position = 0;
            string roundTrip = WordToMarkdownConverter.Convert(ms, new WordToMarkdownOptions());

            Assert.Contains("# Heading 1", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("**world**", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("*universe*", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_Markdown_Lists_RoundTrip() {
            string md = "- Item 1\n- Item 2\n\n1. First\n1. Second";
            using MemoryStream ms = new MemoryStream();
            MarkdownToWordConverter.Convert(md, ms, new MarkdownToWordOptions { FontFamily = "Calibri" });

            ms.Position = 0;
            string roundTrip = WordToMarkdownConverter.Convert(ms, new WordToMarkdownOptions());

            Assert.Contains("- Item 1", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("1. First", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_Markdown_FontResolver() {
            string md = "Hello";
            using MemoryStream ms = new MemoryStream();
            MarkdownToWordConverter.Convert(md, ms, new MarkdownToWordOptions { FontFamily = "monospace" });

            ms.Position = 0;
            using WordprocessingDocument doc = WordprocessingDocument.Open(ms, false);
            RunFonts fonts = doc.MainDocumentPart!.Document.Body!.Descendants<RunFonts>().First();
            Assert.Equal(FontResolver.Resolve("monospace"), fonts.Ascii);
        }
    }
}
