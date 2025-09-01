using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlFootnotes {
        [Fact]
        public void FootnotesRoundTrip() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddFootNote("footnote text");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });

            Assert.Contains("<sup", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("footnote text", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = html.LoadFromHtml();

            Assert.NotNull(roundTrip.FootNotes);
            Assert.True(roundTrip.FootNotes!.Count >= 1);
            var footnote = roundTrip.FootNotes[0];
            Assert.NotNull(footnote);
            Assert.Equal("footnote text", footnote.Paragraphs[1].Text);

            string html2 = roundTrip.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });
            Assert.Contains("footnote text", html2, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FootnotesCanBeOmitted() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddFootNote("hidden");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = false });

            Assert.DoesNotContain("<sup", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("hidden", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}

