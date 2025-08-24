using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Address_ParsesLikeDiv() {
            string html = "<address style=\"text-align:right\"><p>Location</p></address>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("Location", paragraph.Text);
            Assert.Equal(JustificationValues.Right, paragraph.ParagraphAlignment);
        }

        [Fact]
        public void HtmlToWord_Article_ParsesLikeDiv() {
            string html = "<article style=\"text-align:center\"><p>Article text</p></article>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("Article text", paragraph.Text);
            Assert.Equal(JustificationValues.Center, paragraph.ParagraphAlignment);
        }

        [Fact]
        public void HtmlToWord_Aside_ParsesLikeDiv() {
            string html = "<aside style=\"text-align:justify\"><p>Aside note</p></aside>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("Aside note", paragraph.Text);
            Assert.Equal(JustificationValues.Both, paragraph.ParagraphAlignment);
        }

        [Fact]
        public void HtmlToWord_Nav_ParsesLikeDiv() {
            string html = "<nav style=\"margin-left:20pt;padding-left:10pt\"><p>Menu</p></nav>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];

            Assert.Equal("Menu", paragraph.Text);
            Assert.Equal(30d, paragraph.IndentationBeforePoints);
        }

        [Fact]
        public void HtmlStructuralTags_CreateBookmarks() {
            string html = "<section id=\"intro\"><p>Intro</p></section><article id=\"art\"><p>Article</p></article>";
            using var doc = html.LoadFromHtml();
            Assert.Contains(doc.Bookmarks, b => b.Name == "section:intro");
            Assert.Contains(doc.Bookmarks, b => b.Name == "article:art");

            string roundTrip = doc.ToHtml();
            System.Console.WriteLine(roundTrip);
            Assert.Contains("<section id=\"intro\">", roundTrip);
            Assert.Contains("<article id=\"art\">", roundTrip);
        }
    }
}

