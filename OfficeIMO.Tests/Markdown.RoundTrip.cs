using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class MarkdownRoundTripTests : Word {
        [Fact]
        public void Markdown_To_Word_To_Markdown_RoundTrip_Preserves_CoreFeatures() {
            string md = "" +
                "# Report\n" +
                "Intro with **bold**, *italic*, ~~strike~~, <u>underline</u>, and `code`.\\n\n" +
                "- Item 1\n" +
                "- [x] Task done\n" +
                "- [ ] Task todo\n" +
                "- Link: [Docs](https://example.com)\n\n" +
                "1. First\n2. Second\n\n" +
                "| Name | Score | Date |\n" +
                "|:-----|------:|:----:|\n" +
                "| Alice | 98.5 | 2024-01-10 |\n" +
                "| Bob   | 91.0 | 2023-08-22 |\n\n" +
                "Here is a ref[^1].\n\n" +
                "[^1]: Footnote body.";

            using var doc = md.LoadFromMarkdown();
            var md2 = doc.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true });

            Assert.Contains("# Report", md2);
            Assert.Contains("**bold**", md2);
            Assert.Contains("*italic*", md2);
            Assert.Contains("~~strike~~", md2);
            Assert.Contains("<u>underline</u>", md2);
            Assert.Contains("`code`", md2);
            Assert.Contains("- [x] Task done", md2);
            Assert.Contains("- [ ] Task todo", md2);
            Assert.Contains("[Docs](https://example.com)", md2);
            Assert.Contains("| Name | Score | Date |", md2);
            Assert.Contains("|:---", md2); // alignment row present
            Assert.Contains("[^1]", md2);
            Assert.Contains("[^1]:", md2);
        }

        [Fact]
        public void Word_Fluent_To_Markdown_Back_To_Word_Preserves_Structure() {
            using var doc = WordDocument.Create();
            doc.AsFluent()
                .H1("Title")
                .P("Hello world")
                .Ul(ul => ul.Item("One").ItemTask("Done", true).ItemLink("Docs", "https://example.com"))
                .Ol(ol => ol.Item("First").Item("Second"))
                .Paragraph(p => p.Bold("B").Text(" ").Italic("I").Text(" ").Underline("U").Text(" ").Strike("S").Text(" ").Code("C"))
                .Table(t => t.Headers("Name", "Score").Row("Alice", "98.5").Row("Bob", "91.0"));

            var md = doc.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true });
            Assert.Contains("# Title", md);
            Assert.Contains("Hello world", md);
            Assert.Contains("- One", md);
            Assert.Contains("- [x] Done", md);
            Assert.Contains("[Docs](https://example.com)", md);
            Assert.Contains("1. First", md);
            Assert.Contains("**B**", md);
            Assert.Contains("*I*", md);
            Assert.Contains("<u>U</u>", md);
            Assert.Contains("~~S~~", md);
            Assert.Contains("`C`", md);
            Assert.Contains("| Name | Score |", md);

            using var doc2 = md.LoadFromMarkdown();
            // Verify checkbox restored
            var checkboxParagraph = doc2.Paragraphs.FirstOrDefault(p => p.IsCheckBox);
            Assert.NotNull(checkboxParagraph);
            Assert.True(checkboxParagraph!.CheckBox!.IsChecked);
            // Verify heading style
            Assert.True(doc2.Paragraphs.Any(p => p.Style == WordParagraphStyles.Heading1));
        }
    }
}
