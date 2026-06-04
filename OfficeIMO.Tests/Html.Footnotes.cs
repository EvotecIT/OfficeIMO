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
        public void FootnotesRoundTripMultipleParagraphs() {
            using var doc = WordDocument.Create();
            var noteReference = doc.AddParagraph("Hello").AddFootNote("First paragraph");
            noteReference.FootNote!.Paragraphs!.Last().AddParagraph("Second paragraph");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });

            Assert.Contains("<p>First paragraph</p><p>Second paragraph</p>", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = html.LoadFromHtml();

            var footnote = Assert.Single(roundTrip.FootNotes);
            var texts = footnote.Paragraphs!.Skip(1).Select(paragraph => paragraph.Text).ToArray();
            Assert.Equal(new[] { "First paragraph", "Second paragraph" }, texts);
        }

        [Fact]
        public void FootnotesCanBeOmitted() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddFootNote("hidden");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = false });

            Assert.DoesNotContain("<sup", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("hidden", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void EndnotesRoundTrip() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddEndNote("endnote text");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportEndnotes = true });

            Assert.Contains("href=\"#en1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("class=\"endnotes\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("endnote text", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = html.LoadFromHtml();

            var endNotes = roundTrip.EndNotes;
            Assert.NotNull(endNotes);
            var endNote = Assert.Single(endNotes);
            Assert.True(endNote.Paragraphs!.Count > 1);
            Assert.Equal("endnote text", endNote.Paragraphs![1].Text);

            string html2 = roundTrip.ToHtml(new WordToHtmlOptions { ExportEndnotes = true });
            Assert.Contains("class=\"endnotes\"", html2, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("endnote text", html2, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void EndnotesCanBeOmitted() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddEndNote("hidden endnote");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportEndnotes = false });

            Assert.DoesNotContain("href=\"#en", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("class=\"endnotes\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("hidden endnote", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}

