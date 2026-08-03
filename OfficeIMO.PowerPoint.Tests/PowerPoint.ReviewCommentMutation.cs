using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
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
        public void ClassicCommentMutation_RejectsMissingTimestamp() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointClassicComment comment = presentation.AddClassicComment(
                presentation.AddSlide(),
                new PowerPointCommentAuthor("Reviewer", "R"), "Review");
            DateTime? created = comment.Created;

            Assert.Throws<ArgumentNullException>(() => comment.Created = null);

            Assert.Equal(created, comment.Created);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingCommentedSlideReconcilesClassicAndModernAuthors() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide removed = presentation.AddSlide();
            PowerPointSlide retained = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A");
            var bob = new PowerPointCommentAuthor("Bob", "B");
            PowerPointClassicComment removedClassic = presentation.AddClassicComment(
                removed, alice, "Removed first");
            PowerPointClassicComment surviving = presentation.AddClassicComment(
                retained, alice, "Surviving");
            presentation.AddClassicComment(removed, alice, "Removed last");
            PowerPointModernComment modern = presentation.AddModernComment(
                removed, bob, "Modern removed");
            PowerPointModernCommentReply removedReply = modern.AddReply(
                alice, "Modern reply removed");

            presentation.RemoveSlide(0);

            Assert.Throws<InvalidOperationException>(() => removedClassic.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => modern.Status =
                PowerPointModernCommentStatus.Resolved);
            Assert.Throws<InvalidOperationException>(() => removedReply.Text = "Detached");
            Assert.Equal(surviving.Index,
                GetClassicAuthor(presentation, "Alice").LastIndex!.Value);
            Assert.Empty(presentation.OpenXmlDocument.PresentationPart!.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.NotEmpty(presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ModernCommentMutation_PreservesAbsentOptionalAuthorIdentity() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddModernComment(slide,
                new PowerPointCommentAuthor("Imported", "I"), "Review");
            P188.Author author = presentation.OpenXmlDocument.PresentationPart!
                .Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList!.Elements<P188.Author>())
                .Single();
            author.UserId = null;
            author.ProviderId = null;
            author.Initials = null;
            string id = author.Id!.Value!;
            PowerPointModernComment comment = Assert.Single(
                presentation.GetModernComments(slide));

            Assert.Equal("I", comment.Author.Initials);

            comment.SetAuthor(comment.Author);

            P188.Author retained = presentation.OpenXmlDocument.PresentationPart!
                .Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList!.Elements<P188.Author>())
                .Single();
            Assert.Equal(id, retained.Id!.Value);
            Assert.Null(retained.Initials);
            Assert.Null(retained.UserId);
            Assert.Null(retained.ProviderId);
        }

        [Fact]
        public void ClassicCommentMutation_PreservesEmptyImportedInitials() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointClassicComment comment = presentation.AddClassicComment(
                presentation.AddSlide(),
                new PowerPointCommentAuthor("Imported", "I"), "Review");
            P.CommentAuthor stored = GetClassicAuthor(presentation, "Imported");
            stored.Initials = string.Empty;
            uint id = stored.Id!.Value;
            uint index = comment.Index;

            PowerPointCommentAuthor imported = comment.Author;
            Assert.Equal("I", imported.Initials);

            comment.SetAuthor(imported);

            P.CommentAuthor retained = Assert.Single(presentation
                .OpenXmlDocument.PresentationPart!.CommentAuthorsPart!
                .CommentAuthorList!.Elements<P.CommentAuthor>());
            Assert.Equal(id, retained.Id!.Value);
            Assert.Equal(string.Empty, retained.Initials!.Value);
            Assert.Equal(index, comment.Index);
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

        [Fact]
        public void ClassicCommentMutation_ReusesFreeIndexAfterUInt32Maximum() {
            using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            presentation.AddClassicComment(slide, author, "Maximum");
            P.Comment stored = presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().SlideCommentsPart!.CommentList!
                .Elements<P.Comment>().Single();
            stored.Index = uint.MaxValue;
            GetClassicAuthor(presentation, "Reviewer").LastIndex =
                uint.MaxValue;

            presentation.AddClassicComment(slide, author, "Gap");

            Assert.Equal(new[] { 0U, uint.MaxValue }, presentation
                .OpenXmlDocument.PresentationPart!.SlideParts.Single()
                .SlideCommentsPart!.CommentList!.Elements<P.Comment>()
                .Select(comment => comment.Index!.Value)
                .OrderBy(index => index));
            Assert.Equal(uint.MaxValue,
                GetClassicAuthor(presentation, "Reviewer").LastIndex!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClassicCommentMutation_AllocatesWithinBinaryIndexRange() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            presentation.AddClassicComment(slide, author, "Maximum");
            P.Comment maximum = presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().SlideCommentsPart!.CommentList!
                .Elements<P.Comment>().Single();
            maximum.Index = int.MaxValue;
            GetClassicAuthor(presentation, "Reviewer").LastIndex = int.MaxValue;

            presentation.AddClassicComment(slide, author, "Gap");

            Assert.Equal(new[] { 0U, (uint)int.MaxValue }, presentation
                .OpenXmlDocument.PresentationPart!.SlideParts.Single()
                .SlideCommentsPart!.CommentList!.Elements<P.Comment>()
                .Select(comment => comment.Index!.Value)
                .OrderBy(index => index));
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
        }

        [Theory]
        [InlineData("nul\0text")]
        [InlineData("control\u0001text")]
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
        public void ClassicCommentMutation_RejectsBinaryIncompatibleAuthors() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCommentAuthor longName = new(new string('n', 53), "N");
            PowerPointCommentAuthor longInitials = new("Name",
                new string('i', 53));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name\0Suffix", "N"));

            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, longName, "Review"));
            Assert.Throws<ArgumentException>(() =>
                presentation.AddClassicComment(slide, longInitials, "Review"));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, new PowerPointCommentAuthor("Valid", "V"), "Review");
            Assert.Throws<ArgumentException>(() => comment.SetAuthor(longName));
            Assert.Equal("Valid", comment.Author.Name);
        }

        [Fact]
        public void ClassicCommentMutation_RejectsPositionsOutsideBinaryRange() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddClassicComment(slide, author, "Review",
                    (long)int.MaxValue + 1L, 0L));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddClassicComment(slide, author, "Review",
                    0L, (long)int.MinValue - 1L));
            Assert.Empty(presentation.GetClassicComments(slide));

            PowerPointClassicComment comment = presentation.AddClassicComment(
                slide, author, "Valid", 10L, 20L);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.X = long.MaxValue);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.Y = long.MinValue);
            Assert.Equal(10L, comment.X);
            Assert.Equal(20L, comment.Y);
        }

        [Fact]
        public void ModernCommentMutation_RejectsPositionsOutsideDrawingRange() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, author, "Review",
                    x: long.MaxValue));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, author, "Review",
                    y: long.MinValue));
            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.Empty(presentation.GetModernComments(slide));

            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Valid", x: 10L, y: 20L);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.X = long.MaxValue);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.Y = long.MinValue);
            Assert.Equal(10L, comment.X);
            Assert.Equal(20L, comment.Y);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentCoordinateGettersDoNotMaterializeMissingPosition() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Positionless");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            nativeComment.GetFirstChild<P188.Point2DType>()!.Remove();

            Assert.Equal(0L, comment.X);
            Assert.Equal(0L, comment.Y);
            Assert.Null(nativeComment.GetFirstChild<P188.Point2DType>());

            comment.X = 12L;
            Assert.Equal(12L, comment.X);
            Assert.Equal(0L, comment.Y);
            Assert.NotNull(nativeComment.GetFirstChild<P188.Point2DType>());
        }

        [Fact]
        public void ModernReplyMutationRejectsDetachedReplyList() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Review");
            PowerPointModernCommentReply reply = comment.AddReply(author,
                "Reply");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            P188.CommentReplyList replyList = Assert.IsType<P188.CommentReplyList>(
                nativeComment.GetFirstChild<P188.CommentReplyList>());

            replyList.Remove();

            Assert.NotNull(replyList.GetFirstChild<P188.CommentReply>()?.Parent);
            Assert.Throws<InvalidOperationException>(() => reply.Text = "Detached");
            Assert.Throws<InvalidOperationException>(() => reply.Status =
                PowerPointModernCommentStatus.Resolved);
            Assert.Throws<InvalidOperationException>(() =>
                reply.SetAuthor(new PowerPointCommentAuthor("Other", "O")));
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
        public void ModernCommentMutation_RejectsXmlInvalidText() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R",
                "reviewer@example.test", "OfficeIMO");

            Assert.Throws<ArgumentException>(() =>
                presentation.AddModernComment(slide, author, "Bad\0text"));
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, author, "Valid");
            PowerPointModernCommentReply reply = comment.AddReply(author,
                "Valid reply");
            Assert.Throws<ArgumentNullException>(() =>
                comment.Created = null);
            Assert.Throws<ArgumentNullException>(() =>
                reply.Created = null);
            Assert.Throws<ArgumentException>(() => comment.Text = "Bad\u0001text");
            Assert.Throws<ArgumentException>(() => reply.Text = "Bad\u000Btext");
            Assert.Throws<ArgumentException>(() =>
                comment.AddReply(author, "Bad\u000Ctext"));
            Assert.Equal("Valid", comment.Text);
            Assert.Equal("Valid reply", reply.Text);
            Assert.NotNull(comment.Created);
            Assert.NotNull(reply.Created);
            Assert.Single(comment.Replies);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ModernCommentMutation_RejectsInvalidStatusBeforeAuthorCreation() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var alice = new PowerPointCommentAuthor("Alice", "A",
                "alice@example.test", "OfficeIMO");

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddModernComment(slide, alice, "Review",
                    (PowerPointModernCommentStatus)999));

            Assert.Empty(GetModernAuthorNames(presentation));
            Assert.Empty(presentation.GetModernComments(slide));
            Assert.DoesNotContain(presentation.AnalyzeLegacyPptWrite().Findings,
                finding => finding.Code == "PPT-WRITE-MODERN-COMMENTS");

            PowerPointModernComment comment = presentation.AddModernComment(
                slide, alice, "Valid review");
            var bob = new PowerPointCommentAuthor("Bob", "B",
                "bob@example.test", "OfficeIMO");
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                comment.AddReply(bob, "Reply",
                    (PowerPointModernCommentStatus)999));
            Assert.Equal(new[] { "Alice" }, GetModernAuthorNames(presentation));
            Assert.Empty(comment.Replies);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CommentAuthor_RejectsXmlInvalidIdentityFields() {
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Bad\u0001name"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "B\u0001"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "N", "bad\u0001user"));
            Assert.Throws<ArgumentException>(() =>
                new PowerPointCommentAuthor("Name", "N", "user",
                    "bad\u0001provider"));
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
        public void ModernCommentMutation_PreservesProducerMetadataWhenRemovingLastComment() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Review");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            part.CommentList!.Append(new OpenXmlUnknownElement(
                "vendor", "producerData",
                "urn:officeimo:test:comments"));
            string original = part.CommentList.OuterXml;

            Assert.Throws<NotSupportedException>(() => comment.Remove());

            Assert.Equal(original, part.CommentList.OuterXml);
            Assert.Contains(slide.SlidePart.Parts,
                pair => ReferenceEquals(pair.OpenXmlPart, part));
            Assert.Single(presentation.GetModernComments(slide));
            Assert.Equal(new[] { "Reviewer" },
                GetModernAuthorNames(presentation));
        }

        [Fact]
        public void ModernCommentText_RejectsRichBodyReplacementWithoutMutation() {
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
            string original = body.OuterXml;

            Assert.Throws<NotSupportedException>(() =>
                PowerPointPresentation.SetModernCommentText(comment,
                    "Alpha\r\n\r\nBeta"));

            Assert.Equal(original, body.OuterXml);
            Assert.Equal("First\nSecond\nThird",
                PowerPointPresentation.GetModernCommentText(comment));
        }

        [Fact]
        public void RichModernCommentBodyIsPreservedAndReportedAsNonEditable() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointModernComment comment = presentation.AddModernComment(
                slide, new PowerPointCommentAuthor("Reviewer", "R"),
                "Plain");
            PowerPointCommentPart part = Assert.Single(slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>());
            P188.Comment nativeComment = Assert.Single(part.CommentList!
                .Elements<P188.Comment>());
            P188.TextBodyType body = nativeComment
                .GetFirstChild<P188.TextBodyType>()!;
            body.RemoveAllChildren<A.Paragraph>();
            body.Append(new A.Paragraph(
                new A.Run(new A.RunProperties { Bold = true },
                    new A.Text("Rich")),
                new A.Run(new A.RunProperties { Italic = true },
                    new A.Text(" body"))));
            string original = body.OuterXml;

            PowerPointFeatureFinding finding = Assert.Single(presentation
                .InspectFeatures().FindFeatures("Comments"));
            Assert.Equal(PowerPointFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Throws<NotSupportedException>(() =>
                comment.Text = "Replacement");
            Assert.Equal(original, body.OuterXml);
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
