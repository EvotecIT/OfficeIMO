using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlAbbreviations {
        [Fact]
        public void AbbrBecomesFootnoteAndRoundsTrip() {
            const string html = "<abbr title=\"desc\">text</abbr>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
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

        [Fact]
        public void AbbrTitleLinksProtocolRelativeUrl() {
            const string html = "<abbr title=\"//example.com/source\">text</abbr>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            var note = Assert.Single(footNotes!);
            Assert.Contains(note.Paragraphs!, p => p.Hyperlink?.Uri == new Uri("https://example.com/source"));
        }

        [Fact]
        public void AbbrCanUseEndnotes() {
            const string html = "<abbr title=\"desc\">text</abbr>";
            var options = new HtmlToWordOptions { NoteReferenceType = NoteReferenceType.Endnote };
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(options);
            Assert.True(doc.Paragraphs.Count >= 1);
            Assert.Equal("text", doc.Paragraphs[0].Text);
            var endNotes = doc.EndNotes;
            Assert.NotNull(endNotes);
            Assert.Single(endNotes!);
            Assert.Equal("desc", endNotes![0].Paragraphs![1].Text);
        }

        [Fact]
        public void AbbrEndnoteRoundTripsAsAbbrTitle() {
            const string html = "<abbr title=\"desc\">text</abbr>";
            var options = new HtmlToWordOptions { NoteReferenceType = NoteReferenceType.Endnote };
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(options);

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { ExportEndnotes = true });

            Assert.Contains("<abbr", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"desc\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">text</abbr>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<sup", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
