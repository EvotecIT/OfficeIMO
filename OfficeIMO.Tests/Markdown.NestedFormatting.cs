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
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Contains(runs, r => r.Bold && r.Strike && r.Text == "bold strike");
            Assert.Contains(runs, r => r.Bold && r.Italic && r.Text == "bold italic");
        }

        [Fact]
        public void Markdown_NestedFormatting_StrikeUnderline_RoundTrip() {
            string md = "Text ~~<u>strike underline</u>~~ and <u>~~underline strike~~</u>.";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Contains(runs, r => r.Strike && r.Underline == UnderlineValues.Single && r.Text == "strike underline");
            Assert.Contains(runs, r => r.Strike && r.Underline == UnderlineValues.Single && r.Text == "underline strike");

            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true });
            Assert.Contains("~~<u>strike underline</u>~~", roundTrip);
            Assert.Contains("~~<u>underline strike</u>~~", roundTrip);
        }
    }
}
