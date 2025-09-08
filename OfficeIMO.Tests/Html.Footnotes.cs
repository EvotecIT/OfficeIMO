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

            var footNotes = roundTrip.FootNotes;
            Assert.NotNull(footNotes);
            Assert.True(footNotes!.Count >= 1);
            var footnote = footNotes![0];
            Assert.True(footnote.Paragraphs!.Count > 1);
            Assert.Equal("footnote text", footnote.Paragraphs![1].Text);

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

