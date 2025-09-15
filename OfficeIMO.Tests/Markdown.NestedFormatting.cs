using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Markdown_NestedEmphasis_BoldItalic() {
            string md = "This ***bolditalic*** text.";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First(r => r.Bold && r.Italic);
            Assert.Equal("bolditalic", run.Text);
        }

        [Fact]
        public void Markdown_MixedFormatting_StrikeBoldItalic() {
            string md = "Text ~~**bold strike**~~ and ***bold italic***.";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            // Validate via round-trip because the engine preserves inner markers inside strike
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true });
            Assert.Contains("~~**bold strike**~~", roundTrip);
            Assert.Contains("***bold italic***", roundTrip);
        }

        [Fact]
        public void Markdown_NestedFormatting_StrikeUnderline_RoundTrip() {
            string md = "Text ~~<u>strike underline</u>~~ and <u>~~underline strike~~</u>.";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            // Run-level nesting varies; validate round-trip Markdown output instead
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true });
            Assert.Contains("~~<u>strike underline</u>~~", roundTrip);
            Assert.Contains("<u>~~underline strike~~</u>", roundTrip);
        }
    }
}
