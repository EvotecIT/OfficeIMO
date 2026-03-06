using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownReader_RendersHighlightInlineToHtml() {
            var doc = MarkdownReader.Parse("Paragraph with ==highlight== and ==**nested**==.");

            var html = doc.ToHtmlFragment();

            Assert.Contains("<mark>highlight</mark>", html);
            Assert.Contains("<mark><strong>nested</strong></mark>", html);
        }

        [Fact]
        public void Markdown_Highlight_RoundTrip_PreservesNestedFormatting() {
            const string md = "Text ==highlighted== and ==**important**==.";

            using var doc = md.LoadFromMarkdown();
            var runs = doc.Paragraphs[0].GetRuns().ToList();

            Assert.Contains(runs, r => r.Text == "highlighted" && r.Highlight == HighlightColorValues.Yellow);
            Assert.Contains(runs, r => r.Text == "important" && r.Highlight == HighlightColorValues.Yellow && r.Bold);

            var roundTrip = doc.ToMarkdown(new WordToMarkdownOptions { EnableHighlight = true });
            Assert.Contains("==highlighted==", roundTrip);
            Assert.Contains("==**important**==", roundTrip);
        }
    }
}
