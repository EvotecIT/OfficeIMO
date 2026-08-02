using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.Tests {
    public class PowerPointReviewCommentMutationTests {
        [Fact]
        public void ClassicComments_CreateEditReassignRemove_AndRoundTripPpt() {
            byte[] binary;
            using (PowerPointPresentation presentation = PowerPointPresentation.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                var alice = new PowerPointCommentAuthor("Alice Reviewer", "AR");
                var bob = new PowerPointCommentAuthor("Bob Reviewer", "BR");
                PowerPointClassicComment first = presentation.AddClassicComment(slide,
                    alice, "First review", 120, 240,
                    new DateTime(2026, 8, 1, 9, 0, 0, DateTimeKind.Utc));
                PowerPointClassicComment removed = presentation.AddClassicComment(slide,
                    alice, "Remove me");

                first.Text = "Updated review";
                first.X = -25;
                first.Y = 900;
                first.SetAuthor(bob);
                removed.Remove();

                PowerPointClassicComment current = Assert.Single(
                    presentation.GetClassicComments(slide));
                Assert.Equal("Bob Reviewer", current.Author.Name);
                Assert.Equal("Updated review", current.Text);
                Assert.Equal(-25, current.X);
                Assert.Equal(900, current.Y);
                var preflight = presentation.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite, string.Join(Environment.NewLine,
                    preflight.Findings.Select(finding => finding.Code + ": " + finding.Description)));
                binary = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            using PowerPointPresentation roundTrip = PowerPointPresentation.Load(
                new MemoryStream(binary, writable: false));
            PowerPointReviewComment projected = Assert.Single(
                roundTrip.InspectReviewComments().Comments);
            Assert.Equal(PowerPointCommentKind.Classic, projected.Kind);
            Assert.Equal("Bob Reviewer", projected.AuthorName);
            Assert.Equal("Updated review", projected.Text);
            Assert.Equal(-25, projected.X);
            Assert.Equal(900, projected.Y);
        }

        [Fact]
        public void ClassicCommentMutation_RecomputesRetainedAuthorLastIndex() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            presentation.AddClassicComment(slide, alice, "First");
            PowerPointClassicComment second = presentation.AddClassicComment(
                slide, alice, "Second");
            PowerPointClassicComment third = presentation.AddClassicComment(
                slide, alice, "Third");

            third.SetAuthor(bob);
            Assert.Equal(2U, GetClassicAuthor(presentation, "Alice").LastIndex!.Value);

            second.Remove();
            Assert.Equal(1U, GetClassicAuthor(presentation, "Alice").LastIndex!.Value);
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.NotEmpty(presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ClassicCommentMutation_ReusesFreeAuthorIdWhenMaximumIsOccupied() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            CommentAuthorsPart authorsPart = presentation.OpenXmlDocument
                .PresentationPart!.AddNewPart<CommentAuthorsPart>();
            authorsPart.CommentAuthorList = new P.CommentAuthorList(
                new P.CommentAuthor {
                    Id = uint.MaxValue,
                    Name = "Maximum",
                    Initials = "M",
                    LastIndex = 0U,
                    ColorIndex = 0U
                });

            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Available", "A"), "Review");

            Assert.Equal(0U, GetClassicAuthor(presentation, "Available").Id!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Theory]
        [InlineData("nul\0text")]
        public void ClassicCommentMutation_RejectsBinaryIncompatibleText(
            string invalidText) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, author, invalidText));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid");
            Assert.Throws<ArgumentException>(() => comment.Text = invalidText);
            Assert.Equal("Valid", comment.Text);
        }

        [Fact]
        public void ClassicCommentMutation_RejectsTextAboveBinaryLimit() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            string invalidText = new string('x', 32001);

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, author, invalidText));
            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid");
            Assert.Throws<ArgumentException>(() => comment.Text = invalidText);
            Assert.Equal("Valid", comment.Text);
        }

        [Fact]
        public void ModernComments_CreateEditReplyReassignRemove_AndRoundTripPptx() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                var alice = new PowerPointCommentAuthor("Alice Reviewer", "AR",
                    "alice@example.test", "OfficeIMO");
                var bob = new PowerPointCommentAuthor("Bob Reviewer", "BR",
                    "bob@example.test", "OfficeIMO");
                PowerPointModernComment comment = presentation.AddModernComment(slide,
                    alice, "Modern review", PowerPointModernCommentStatus.Active,
                    123, 456, new DateTime(2026, 8, 1, 10, 0, 0, DateTimeKind.Utc));
                PowerPointModernCommentReply reply = comment.AddReply(bob, "Initial reply");
                PowerPointModernCommentReply removed = comment.AddReply(alice, "Remove reply");

                comment.Text = "Updated modern review";
                comment.Status = PowerPointModernCommentStatus.Resolved;
                comment.X = 321;
                comment.SetAuthor(bob);
                reply.Text = "Updated reply";
                reply.Status = PowerPointModernCommentStatus.Closed;
                removed.Remove();

                Assert.Equal("Bob Reviewer", comment.Author.Name);
                Assert.Equal("Bob Reviewer", reply.Author.Name);
                Assert.Single(comment.Replies);
                Assert.Contains(presentation.AnalyzeLegacyPptWrite().Findings,
                    finding => finding.Code == "PPT-WRITE-MODERN-COMMENTS");
                PowerPointFeatureFinding finding = presentation.InspectFeatures().Features
                    .Single(item => item.Name == "Comments");
                Assert.Equal(PowerPointFeatureSupportLevel.Editable, finding.SupportLevel);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointModernComment current = Assert.Single(
                reopened.GetModernComments(reopened.Slides[0]));
            Assert.Equal("Updated modern review", current.Text);
            Assert.Equal(PowerPointModernCommentStatus.Resolved, current.Status);
            Assert.Equal(321, current.X);
            PowerPointModernCommentReply currentReply = Assert.Single(current.Replies);
            Assert.Equal("Updated reply", currentReply.Text);
            Assert.Equal(PowerPointModernCommentStatus.Closed, currentReply.Status);
            PowerPointReviewReport report = reopened.InspectReviewComments();
            Assert.Equal(2, report.ModernCount);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));

            current.Remove();
            Assert.Empty(reopened.GetModernComments(reopened.Slides[0]));
        }

        [Fact]
        public void CommentMutation_RejectsForeignSlidesAndEmptyText() {
            using PowerPointPresentation first = PowerPointPresentation.Create();
            using PowerPointPresentation second = PowerPointPresentation.Create();
            PowerPointSlide foreign = second.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer");

            Assert.Throws<ArgumentException>(() =>
                first.AddClassicComment(foreign, author, "Review"));
            Assert.Throws<ArgumentException>(() =>
                first.AddModernComment(foreign, author, "Review"));
            Assert.Throws<ArgumentException>(() =>
                first.AddClassicComment(first.AddSlide(), author, " "));
        }

        [Fact]
        public void CommentAuthor_GeneratesUnicodeScalarInitialsThatRoundTrip() {
            const string supplementaryCjk = "\U00020000";
            var emojiAuthor = new PowerPointCommentAuthor("😀 Reviewer");
            var cjkAuthor = new PowerPointCommentAuthor(supplementaryCjk + " Reviewer");

            Assert.Equal("😀R", emojiAuthor.Initials);
            Assert.Equal(supplementaryCjk + "R", cjkAuthor.Initials);

            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                presentation.AddClassicComment(slide, emojiAuthor, "Emoji author");
                presentation.AddModernComment(slide, cjkAuthor, "CJK author");
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.Equal("😀R", Assert.Single(reopened.GetClassicComments(
                reopened.Slides[0])).Author.Initials);
            Assert.Equal(supplementaryCjk + "R", Assert.Single(
                reopened.GetModernComments(reopened.Slides[0])).Author.Initials);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ModernCommentMutation_RemovesUnusedAuthorMetadata() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, alice, "Review");
            PowerPointModernCommentReply reply = comment.AddReply(alice, "Reply");

            comment.SetAuthor(bob);
            Assert.Equal(new[] { "Alice", "Bob" }, GetModernAuthorNames(presentation));
            reply.SetAuthor(bob);
            Assert.Equal(new[] { "Bob" }, GetModernAuthorNames(presentation));

            comment.Remove();
            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.DoesNotContain(presentation.OpenXmlDocument.PresentationPart!.Parts,
                pair => pair.OpenXmlPart is PowerPointAuthorsPart);
            Assert.Empty(presentation.GetModernComments(slide));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentText_PreservesParagraphsBreaksAndBlankLines() {
            var comment = new P188.Comment();
            var body = new P188.TextBodyType(new A.BodyProperties(),
                new A.ListStyle());
            body.Append(
                new A.Paragraph(new A.Run(new A.Text("First")),
                    new A.Break(), new A.Run(new A.Text("Second"))),
                new A.Paragraph(new A.Run(new A.Text("Third"))));
            comment.Append(body);

            Assert.Equal("First\nSecond\nThird",
                PowerPointPresentation.GetModernCommentText(comment));

            PowerPointPresentation.SetModernCommentText(comment,
                "Alpha\r\n\r\nBeta");

            Assert.Equal(3, body.Elements<A.Paragraph>().Count());
            Assert.Equal("Alpha\n\nBeta",
                PowerPointPresentation.GetModernCommentText(comment));
        }

        private static string[] GetModernAuthorNames(
            PowerPointPresentation presentation) => presentation.OpenXmlDocument
                .PresentationPart!.Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList?.Elements<P188.Author>()
                    ?? Enumerable.Empty<P188.Author>())
                .Select(author => author.Name?.Value ?? string.Empty)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

        private static P.CommentAuthor GetClassicAuthor(
            PowerPointPresentation presentation, string name) => presentation
                .OpenXmlDocument.PresentationPart!.CommentAuthorsPart!
                .CommentAuthorList!.Elements<P.CommentAuthor>()
                .Single(author => author.Name?.Value == name);
    }
}
