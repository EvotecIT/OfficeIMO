using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.Tests {
    public class PowerPointModernCommentRelationshipTests {
        private const string CommentExtensionUri = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}";

        [Fact]
        public void ModernComments_SaveExplicitSlideReferenceAndPreserveOtherExtensions() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var unrelated = new P.SlideExtension(
                new OpenXmlUnknownElement("e", "producer", "urn:producer")
            ) { Uri = "{0159D55F-9220-4A30-8126-106DAFC44E36}" };
            slide.SlidePart.Slide.Append(new P.SlideExtensionList(unrelated));
            string metadata = unrelated.OuterXml;
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            PowerPointModernComment first = presentation.AddModernComment(slide, author, "First");
            first.AddReply(author, "Reply");
            presentation.AddModernComment(slide, author, "Second");

            AssertSlideReference(slide);
            Assert.Equal(metadata, unrelated.OuterXml);
            using var artifact = PresentationDocument.Open(new MemoryStream(presentation.ToBytes(PowerPointFileFormat.Pptx)), false);
            SlidePart savedSlide = artifact.PresentationPart!.SlideParts.Single();
            AssertReference(savedSlide);
            Assert.Equal(metadata, savedSlide.Slide.SlideExtensionList!.Elements<P.SlideExtension>()
                .Single(extension => extension.Uri!.Value == unrelated.Uri!.Value).OuterXml);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(artifact));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RemovingLastModernCommentRemovesOnlyItsSlideReferenceAndAllowsRecreation(bool withOtherExtension) {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var author = new PowerPointCommentAuthor("Reviewer", "R");
            PowerPointModernComment first = presentation.AddModernComment(slide, author, "First");
            PowerPointModernComment second = presentation.AddModernComment(slide, author, "Second");
            // A valid producer reference also exercises removal on imported documents.
            PowerPointCommentPart part = CommentPart(slide.SlidePart);
            if (slide.SlidePart.Slide.SlideExtensionList == null) {
                slide.SlidePart.Slide.Append(new P.SlideExtensionList(new P.SlideExtension(
                    new P188.CommentRelationship { Id = slide.SlidePart.GetIdOfPart(part) }) { Uri = CommentExtensionUri }));
            }
            P.SlideExtension? unrelated = null;
            if (withOtherExtension) {
                unrelated = new P.SlideExtension(new OpenXmlUnknownElement("e", "producer", "urn:producer")) {
                    Uri = "{0159D55F-9220-4A30-8126-106DAFC44E36}" };
                slide.SlidePart.Slide.SlideExtensionList!.Append(unrelated);
            }
            first.Remove();
            AssertSlideReference(slide);
            second.Remove();
            Assert.Empty(slide.SlidePart.Parts.Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>());
            Assert.Empty(slide.SlidePart.Slide.Descendants<P188.CommentRelationship>());
            if (withOtherExtension) Assert.Same(unrelated, Assert.Single(slide.SlidePart.Slide.SlideExtensionList!.ChildElements));
            else Assert.Null(slide.SlidePart.Slide.SlideExtensionList);
            presentation.AddModernComment(slide, author, "Recreated");
            AssertSlideReference(slide);
            using var artifact = PresentationDocument.Open(new MemoryStream(presentation.ToBytes(PowerPointFileFormat.Pptx)), false);
            AssertReference(artifact.PresentationPart!.SlideParts.Single());
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2021).Validate(artifact));
        }

        [Fact]
        public void AppendingToLoadedUnreferencedCommentPartRestoresItsSlideReference() {
            byte[] legacy;
            using (PowerPointPresentation presentation = PowerPointPresentation.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                presentation.AddModernComment(slide, new PowerPointCommentAuthor("Reviewer", "R"), "Legacy root");
                slide.SlidePart.Slide.SlideExtensionList?.Remove();
                legacy = presentation.ToBytes(PowerPointFileFormat.Pptx);
            }
            using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(legacy));
            PowerPointSlide target = reopened.Slides.Single();
            PowerPointCommentPart original = CommentPart(target.SlidePart);
            reopened.AddModernComment(target, new PowerPointCommentAuthor("Reviewer", "R"), "New root");
            Assert.Same(original, CommentPart(target.SlidePart));
            Assert.Equal(new[] { "Legacy root", "New root" }, reopened.GetModernComments(target).Select(comment => comment.Text));
            AssertSlideReference(target);
            using var artifact = PresentationDocument.Open(new MemoryStream(reopened.ToBytes(PowerPointFileFormat.Pptx)), false);
            AssertReference(artifact.PresentationPart!.SlideParts.Single());
        }

        private static PowerPointCommentPart CommentPart(SlidePart slide) => Assert.Single(slide.Parts
            .Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>());

        private static void AssertSlideReference(PowerPointSlide slide) => AssertReference(slide.SlidePart);

        private static void AssertReference(SlidePart slide) {
            PowerPointCommentPart part = CommentPart(slide);
            P.SlideExtension extension = Assert.Single(slide.Slide.SlideExtensionList!
                .Elements<P.SlideExtension>(), item => item.Uri!.Value == CommentExtensionUri);
            P188.CommentRelationship reference = Assert.Single(extension.Elements<P188.CommentRelationship>());
            Assert.Equal(slide.GetIdOfPart(part), reference.Id!.Value);
            Assert.Same(part, slide.GetPartById(reference.Id!.Value!));
        }
    }
}
