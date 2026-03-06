using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_ParsesHeadingBoldAndLink() {
            string md = "# Heading 1\n\nThis is **bold** with a [link](https://example.com).";

            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs.First(p => p.Style == WordParagraphStyles.Heading1).Style);
            var bodyParagraph = doc.Paragraphs.First(p => p.Text.Contains("bold"));
            var boldRun = bodyParagraph.GetRuns().First(r => r.Bold);
            Assert.Equal("bold", boldRun.Text);
            Assert.Single(doc.HyperLinks);
            Assert.Equal(new Uri("https://example.com"), doc.HyperLinks[0].Uri);
        }

        [Fact]
        public void MarkdownToWord_PreservesFormattingInsideLinkLabels() {
            string md = "This is [**bold link**](https://example.com) and [==highlighted==](https://example.org).";

            using var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var hyperlinkRuns = doc.Paragraphs[0].GetRuns().Where(r => r.IsHyperLink).ToList();

            Assert.Contains(hyperlinkRuns, r =>
                string.Equals(r.Text, "bold link", StringComparison.Ordinal) &&
                r.Bold &&
                string.Equals(r.Hyperlink?.Uri?.ToString(), "https://example.com/", StringComparison.Ordinal));

            Assert.Contains(hyperlinkRuns, r =>
                string.Equals(r.Text, "highlighted", StringComparison.Ordinal) &&
                r.Highlight == DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Yellow &&
                string.Equals(r.Hyperlink?.Uri?.ToString(), "https://example.org/", StringComparison.Ordinal));

            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions { EnableHighlight = true });

            Assert.Matches("\\[\\*\\*bold link\\*\\*\\]\\(https://example\\.com/?\\)", roundTrip);
            Assert.Matches("\\[==highlighted==\\]\\(https://example\\.org/?\\)", roundTrip);
        }
    }
}
