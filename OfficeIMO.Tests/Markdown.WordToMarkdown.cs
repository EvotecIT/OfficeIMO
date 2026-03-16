using System;
using System.IO;
using OfficeIMO.Markdown;
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
            Assert.Matches("\\[OfficeIMO\\]\\(https://example\\.com/?\\)", markdown);
            Assert.Contains("| H1 | H2 |", markdown);
            Assert.Contains("data:image/png;base64", markdown);
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

        [Fact]
        public void WordToMarkdown_ExportsImagesToFiles() {
            using var doc = WordDocument.Create();
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);

            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            var options = new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.File,
                ImageDirectory = tempDir
            };

            string markdown = doc.ToMarkdown(options);

            string fileName = Path.GetFileName(imagePath);
            Assert.Contains($"![", markdown);
            Assert.Contains(fileName, markdown);
            Assert.True(File.Exists(Path.Combine(tempDir, fileName)));
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_BuildsTypedAstDirectly() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Water");
            paragraph.AddText("2").SetVerticalTextAlignment(DocumentFormat.OpenXml.Wordprocessing.VerticalPositionValues.Subscript);
            paragraph.AddText(" and ");
            paragraph.AddText("important").Underline = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single;

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions { EnableUnderline = true });
            var block = Assert.IsType<ParagraphBlock>(Assert.Single(markdown.Blocks));

            Assert.Contains(block.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "sub");
            Assert.Contains(block.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "u");

            string renderedHtml = markdown.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<sub>2</sub>", renderedHtml, StringComparison.Ordinal);
            Assert.Contains("<u>important</u>", renderedHtml, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_Preserves_NonParagraph_Footnote_Blocks() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Body").AddFootNote("Console.WriteLine(1);");
            var footnoteParagraph = doc.FootNotes[0].Paragraphs![1];
            footnoteParagraph.SetStyleId("CodeLang_csharp");

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());
            var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(markdown.Blocks, block => block is FootnoteDefinitionBlock));
            var code = Assert.IsType<CodeBlock>(Assert.Single(footnote.Blocks));

            Assert.Equal("csharp", code.Language);
            Assert.Equal("Console.WriteLine(1);", code.Content);
            Assert.Empty(footnote.ParagraphBlocks);

            string renderedHtml = markdown.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li id=\"fn:1\"><pre><code class=\"language-csharp\">Console.WriteLine(1);", renderedHtml, StringComparison.Ordinal);
            Assert.Contains("<a class=\"footnote-backref\"", renderedHtml, StringComparison.Ordinal);
        }
    }
}

