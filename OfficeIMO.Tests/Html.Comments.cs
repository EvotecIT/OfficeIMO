using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlComments {
        [Fact]
        public void WordToHtml_Comments_AreOptIn() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Commented paragraph");
            paragraph.AddComment("Jane Doe", "JD", "Review note");

            string html = doc.ToHtml();

            Assert.DoesNotContain("class=\"comments\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Review note", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_Comments_ExportLinkedReferencesAndReplies() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Commented paragraph");
            paragraph.AddComment("Jane Doe", "JD", "Review note");
            var comment = Assert.Single(doc.Comments);
            comment.AddReply("Alex Roe", "AR", "Reply note");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportComments = true });

            Assert.Contains("href=\"#comment1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("id=\"commentref1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("class=\"comments\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-author=\"Jane Doe\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-initials=\"JD\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Review note", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("class=\"comment-replies\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("id=\"comment1-reply1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-author=\"Alex Roe\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Reply note", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_Comments_RoundTripAsNativeComments() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Commented paragraph");
            paragraph.AddComment("Jane Doe", "JD", "Review note");
            var comment = Assert.Single(doc.Comments);
            comment.AddReply("Alex Roe", "AR", "Reply note");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportComments = true });

            using var roundTrip = html.LoadFromHtml();

            Assert.DoesNotContain(roundTrip.Paragraphs.Select(paragraph => paragraph.Text), text => text.Contains("Review note", StringComparison.OrdinalIgnoreCase));
            var rootComment = Assert.Single(roundTrip.Comments, comment => string.IsNullOrEmpty(comment.ParentParaId));
            Assert.Equal("Jane Doe", rootComment.Author);
            Assert.Equal("JD", rootComment.Initials);
            Assert.Equal("Review note", rootComment.Text);
            var reply = Assert.Single(rootComment.Replies);
            Assert.Equal("Alex Roe", reply.Author);
            Assert.Equal("AR", reply.Initials);
            Assert.Equal("Reply note", reply.Text);
        }
    }
}
