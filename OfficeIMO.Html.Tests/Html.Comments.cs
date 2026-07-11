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

            using var roundTrip = html.ToWordDocument();

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

        [Fact]
        public void WordToHtml_Comments_RoundTripPreservesDates() {
            var commentDate = new DateTime(2026, 6, 4, 12, 34, 56, DateTimeKind.Utc);
            var replyDate = new DateTime(2026, 6, 4, 13, 15, 0, DateTimeKind.Utc);
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Commented paragraph");
            paragraph.AddComment("Jane Doe", "JD", "Review note");
            var comment = Assert.Single(doc.Comments);
            comment.DateTime = commentDate;
            var reply = comment.AddReply("Alex Roe", "AR", "Reply note");
            reply.DateTime = replyDate;

            string html = doc.ToHtml(new WordToHtmlOptions { ExportComments = true });

            using var roundTrip = html.ToWordDocument();

            var rootComment = Assert.Single(roundTrip.Comments, comment => string.IsNullOrEmpty(comment.ParentParaId));
            var rootReply = Assert.Single(rootComment.Replies);
            Assert.Equal(commentDate, rootComment.DateTime);
            Assert.Equal(replyDate, rootReply.DateTime);
        }

        [Fact]
        public void HtmlToWord_RawHtmlComments_CanImportAsNativeWordComments() {
            var options = new HtmlToWordOptions {
                ImportHtmlComments = true,
                HtmlCommentAuthor = "HTML Reviewer",
                HtmlCommentInitials = "HR"
            };
            string html = "<p>Visible <!-- reviewer note -->text.</p>";

            using var doc = html.ToWordDocument(options);

            Assert.Equal("Visible text.", string.Concat(doc.Paragraphs.Select(paragraph => paragraph.Text)));
            Assert.DoesNotContain(options.Diagnostics, diagnostic => diagnostic.Code == "HtmlCommentSkipped");
            var comment = Assert.Single(doc.Comments);
            Assert.Equal("HTML Reviewer", comment.Author);
            Assert.Equal("HR", comment.Initials);
            Assert.Equal("reviewer note", comment.Text);

            using var stream = doc.SaveAsMemoryStream();
            using var loaded = WordDocument.Load(stream);
            var loadedComment = Assert.Single(loaded.Comments);
            Assert.Equal("HTML Reviewer", loadedComment.Author);
            Assert.Equal("HR", loadedComment.Initials);
            Assert.Equal("reviewer note", loadedComment.Text);

            string roundTrip = loaded.ToHtml(new WordToHtmlOptions { ExportComments = true });
            Assert.Contains("class=\"comments\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-author=\"HTML Reviewer\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-initials=\"HR\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("reviewer note", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
