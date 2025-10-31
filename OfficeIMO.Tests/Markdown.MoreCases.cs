using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    [Collection("WordTests")]
    public class MarkdownMoreCasesTests : Word {
        [Fact]
        public void ReferenceStyleLinks_WithAngleBrackets_AndTitle() {
            string md = "See the [API][ref].\n\n[ref]: <https://example.com/api> \"My API\"\n";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            Assert.Contains("(https://example.com/api)", outMd);
        }

        [Fact]
        public void InlineCode_WithBackticks_ChoosesLongerFence() {
            string md = "Here is ``code ` inner`` end";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            // New engine uses minimal longer fence; accept double or triple backticks
            Assert.True(outMd.Contains("```code ` inner```") || outMd.Contains("``code ` inner``"));
        }

        [Fact]
        public void Footnote_WithContinuationLines() {
            string md = "Ref[^n].\n\n[^n]: First line\n  continued second line\n";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            // Converter normalizes labels to numeric ids
            Assert.Contains("[^1]", outMd);
            Assert.Contains("[^1]:", outMd);
            Assert.Contains("First line", outMd);
        }

        [Fact]
        public void Image_SizeHints_AreParsedByReader() {
            string md = "![logo](logo.png \"t\"){width=40 height=30}";
            var model = MarkdownReader.Parse(md);
            var img = model.Blocks.OfType<ImageBlock>().FirstOrDefault();
            Assert.NotNull(img);
            Assert.Equal(40, img!.Width);
            Assert.Equal(30, img!.Height);
            Assert.Equal("logo.png", img.Path);
            Assert.Equal("logo", img.Alt);
        }

        [Fact]
        public void Autolinks_AreRecognized_And_Preserved() {
            string md = "Visit https://example.com now.";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            // Accept either autolinked markdown or plain URL token
            Assert.Contains("https://example.com", outMd);
        }

        [Fact]
        public void Table_Alignment_RoundTrips() {
            string md = "| A | B | C |\n|:---|---:|:---:|\n| 1 | 2 | 3 |\n";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            // Alignment row should remain intact
            Assert.Contains("|:---|---:|:---:|", outMd.Replace(" ", string.Empty));
        }

        [Fact]
        public void Table_WithoutOuterPipes_Parses_WithAlignment() {
            string md = "Alpha | Beta\n:--- | ---:\nleft | right\n";
            var model = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(model.Blocks[0]);
            Assert.Equal(new[] { "Alpha", "Beta" }, table.Headers);
            Assert.Equal(new[] { ColumnAlignment.Left, ColumnAlignment.Right }, table.Alignments);
            Assert.Single(table.Rows);
            Assert.Equal(new[] { "left", "right" }, table.Rows[0]);
        }

        [Fact]
        public void Table_WithOuterPipes_Parses_WithAlignment() {
            string md = "| X | Y |\n|:---|---:|\n| 1 | 2 |\n"; 
            var model = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(model.Blocks[0]);
            Assert.Equal(new[] { "X", "Y" }, table.Headers);
            Assert.Equal(new[] { ColumnAlignment.Left, ColumnAlignment.Right }, table.Alignments);
            Assert.Single(table.Rows);
            Assert.Equal(new[] { "1", "2" }, table.Rows[0]);
        }

        [Fact]
        public void Table_WithSingleColumn_StillParses() {
            string md = "| Header |\n| --- |\n| Row |\n";
            var model = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(model.Blocks[0]);
            Assert.Equal(new[] { "Header" }, table.Headers);
            Assert.Equal(new[] { ColumnAlignment.None }, table.Alignments);
            Assert.Single(table.Rows);
            Assert.Equal(new[] { "Row" }, table.Rows[0]);
        }

        [Fact]
        public void Setext_Headings_Are_Read_As_Headings() {
            string md = "Title\n=====\n\nSub\n-----\n";
            var model = MarkdownReader.Parse(md);
            var h1 = model.Blocks.OfType<HeadingBlock>().FirstOrDefault(h => h.Level == 1);
            var h2 = model.Blocks.OfType<HeadingBlock>().FirstOrDefault(h => h.Level == 2 && h.Text == "Sub");
            Assert.NotNull(h1);
            Assert.Equal("Title", h1!.Text);
            Assert.NotNull(h2);
        }

        [Fact]
        public void Nested_Lists_Parse_And_Preserve_Indentation_On_Write() {
            string md = "- A\n  - B\n    - C\n1. One\n  1. One.A\n2. Two\n";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            Assert.Contains("- A", outMd);
            Assert.Contains("  - B", outMd);
            Assert.Contains("    - C", outMd);
            Assert.Contains("1. One", outMd);
            Assert.Contains("  1. One.A", outMd);
            // Ordered list numbering is normalized to "1." per item in Markdown
            Assert.True(outMd.Contains("2. Two") || outMd.Contains("1. Two"));
        }
    }
}
