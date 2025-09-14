using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
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
            // Expect at least triple backticks fence to accommodate the backtick inside
            Assert.Contains("```code ` inner```", outMd);
        }

        [Fact]
        public void Footnote_WithContinuationLines() {
            string md = "Ref[^n].\n\n[^n]: First line\n  continued second line\n";
            using var doc = md.LoadFromMarkdown();
            var outMd = doc.ToMarkdown();
            Assert.Contains("[^n]", outMd);
            Assert.Contains("[^n]: First line continued second line", outMd);
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
            // We expect a link markdown token with the same URL as text and href
            Assert.Contains("[https://example.com](https://example.com)", outMd);
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
            Assert.Contains("2. Two", outMd);
        }
    }
}
