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
    }
}
