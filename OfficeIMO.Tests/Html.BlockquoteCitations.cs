using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlBlockquoteCitations {
        [Fact]
        public void BlockquoteWithCitationCreatesFootnote() {
            string html = "<blockquote cite=\"https://example.com\"><p>First</p><p>Second</p></blockquote>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Equal("First", doc.Paragraphs[0].Text);
            Assert.Contains(doc.Paragraphs, p => p.Text == "Second");
            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            Assert.Single(footNotes!);
            Assert.Contains(footNotes![0].Paragraphs!, p => p.Hyperlink?.Uri == new Uri("https://example.com"));
        }

        [Fact]
        public void BlockquoteCitationDoesNotLinkRejectedFileUrl() {
            string html = "<blockquote cite=\"file:///C:/temp/doc.txt\"><p>Quoted</p></blockquote>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var footNotes = doc.FootNotes;
            Assert.NotNull(footNotes);
            var note = Assert.Single(footNotes!);
            Assert.DoesNotContain(note.Paragraphs!, p => p.Hyperlink?.Uri?.Scheme == Uri.UriSchemeFile);
            Assert.Contains(note.Paragraphs!, p => p.Text.Contains("file:///C:/temp/doc.txt", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void BlockquoteCitationRoundTripsAsCiteAttribute() {
            string html = "<blockquote cite=\"https://example.com/source\"><p>Quoted text</p></blockquote>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });

            Assert.Contains("<blockquote cite=\"https://example.com/source\">", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Quoted text", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<sup", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("class=\"footnotes\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void BlockquoteCitationSurvivesSavedDocumentRoundTrip() {
            string html = "<blockquote cite=\"https://example.com/persisted\"><p>Persisted quote</p></blockquote>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            using MemoryStream packageStream = doc.SaveAsMemoryStream();
            using var loaded = WordDocument.Load(packageStream);

            string roundTrip = loaded.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });

            Assert.Contains("<blockquote cite=\"https://example.com/persisted\">", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Persisted quote", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<sup", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("class=\"footnotes\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
