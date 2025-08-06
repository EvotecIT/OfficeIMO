using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_ConvertsElements() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;

            var paragraph = doc.AddParagraph("This is ");
            paragraph.AddText("bold").Bold = true;
            paragraph.AddText(" and ");
            paragraph.AddText("italic").Italic = true;
            paragraph.AddText(" with ");
            paragraph.AddText("strike").Strike = true;
            paragraph.AddText(" and ");
            paragraph.AddText("code").SetFontFamily(FontResolver.Resolve("monospace")!);

            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("Item 1");
            list.AddItem("Item 2");

            var linkParagraph = doc.AddParagraph("Visit ");
            linkParagraph.AddHyperLink("OfficeIMO", new Uri("https://example.com"));

            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "C1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "C2";

            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("# Heading", markdown);
            Assert.Contains("**bold**", markdown);
            Assert.Contains("*italic*", markdown);
            Assert.Contains("~~strike~~", markdown);
            Assert.Contains("`code`", markdown);
            Assert.Contains("- Item 1", markdown);
            Assert.Contains("[OfficeIMO](https://example.com/)", markdown);
            Assert.Contains("| H1 | H2 |", markdown);
            Assert.Contains("![", markdown);
        }

        [Fact]
        public void WordToMarkdown_HandlesFootNotes() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddFootNote("First note");
            doc.AddParagraph("World").AddFootNote("Second note");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("Hello[^1]", markdown);
            Assert.Contains("World[^2]", markdown);
            Assert.Contains("[^1]: First note", markdown);
            Assert.Contains("[^2]: Second note", markdown);
        }
    }
}

