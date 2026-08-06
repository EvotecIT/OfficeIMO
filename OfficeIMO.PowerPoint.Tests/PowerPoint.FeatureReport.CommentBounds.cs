using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointFeatureReportCommentBoundsTests {
        [Fact]
        public void FeatureReportPreservesClassicCommentsOutsideBinaryBounds() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Reviewer", "R"), "Review");
            Comment comment = presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Single().SlideCommentsPart!.CommentList!
                .Elements<Comment>().Single();

            comment.Text!.Text = new string('x', 32001);
            AssertPreserved(presentation);

            comment.Text.Text = "Review";
            comment.Position!.X = (long)int.MaxValue + 1L;
            AssertPreserved(presentation);

            comment.Position.X = 0L;
            comment.Position.Y = (long)int.MinValue - 1L;
            AssertPreserved(presentation);
        }

        [Fact]
        public void FeatureReportPreservesBinaryIncompatibleClassicAuthorIdentity() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Reviewer", "R"), "Review");
            CommentAuthor author = presentation.OpenXmlDocument
                .PresentationPart!.CommentAuthorsPart!.CommentAuthorList!
                .Elements<CommentAuthor>().Single();

            author.Name = string.Empty;
            AssertPreserved(presentation);

            author.Name = new string('n', 53);
            AssertPreserved(presentation);

            author.Name = "Reviewer";
            author.Initials = new string('i', 53);
            AssertPreserved(presentation);
        }

        [Fact]
        public void FeatureReportPreservesDuplicateClassicAuthorIdentity() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Alice", "A"), "First");
            presentation.AddClassicComment(slide,
                new PowerPointCommentAuthor("Bob", "B"), "Second");
            CommentAuthor[] authors = presentation.OpenXmlDocument
                .PresentationPart!.CommentAuthorsPart!.CommentAuthorList!
                .Elements<CommentAuthor>().OrderBy(author => author.Id!.Value)
                .ToArray();

            authors[1].Name = authors[0].Name!.Value;
            authors[1].Initials = authors[0].Initials!.Value;

            AssertPreserved(presentation);
        }

        private static void AssertPreserved(
            PowerPointPresentation presentation) {
            PowerPointFeatureReport report = presentation.InspectFeatures();
            PowerPointFeatureFinding comments = Assert.Single(
                report.FindFeatures("Comments"));
            Assert.Equal(OfficeFeatureSupportLevel.Preserved,
                comments.SupportLevel);
            Assert.Throws<InvalidOperationException>(() =>
                report.EnsureNoAdvancedFeatures());
        }
    }
}
