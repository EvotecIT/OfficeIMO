using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlAbbreviations {
        [Fact]
        public void AbbrBecomesFootnoteAndRoundsTrip() {
            const string html = "<abbr title=\"desc\">text</abbr>";
            using var doc = html.LoadFromHtml();
            Assert.True(doc.Paragraphs.Count >= 1);
            Assert.Equal("text", doc.Paragraphs[0].Text);
            Assert.NotNull(doc.FootNotes);
            Assert.Single(doc.FootNotes);
            Assert.Equal("desc", doc.FootNotes![0].Paragraphs[1].Text);

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });
            Assert.Contains("<abbr", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"desc\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">text</abbr>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<sup", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
