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
            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            Assert.Single(footNotes!);
            Assert.Equal("desc", footNotes![0].Paragraphs![1].Text);

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });
            Assert.Contains("<abbr", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"desc\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">text</abbr>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<sup", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
