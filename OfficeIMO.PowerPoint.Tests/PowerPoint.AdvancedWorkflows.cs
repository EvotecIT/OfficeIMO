using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests {
    public class PowerPointAdvancedWorkflowTests {
        [Fact]
        public void ReviewAndAnimationInspectionProjectsClassicModernAndTimingMetadata() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox target = slide.AddTextBoxPoints("Animated review target", 40, 40, 240, 50);

            PresentationPart presentationPart = slide.SlidePart.GetParentParts().OfType<PresentationPart>().Single();
            CommentAuthorsPart classicAuthors = presentationPart.AddNewPart<CommentAuthorsPart>();
            classicAuthors.CommentAuthorList = new CommentAuthorList(
                new CommentAuthor {
                    Id = 0U, Name = "Classic Reviewer", Initials = "CR", LastIndex = 1U, ColorIndex = 0U
                });
            SlideCommentsPart classicPart = slide.SlidePart.AddNewPart<SlideCommentsPart>();
            classicPart.CommentList = new CommentList(
                new Comment(
                    new Position { X = 120, Y = 240 },
                    new DocumentFormat.OpenXml.Presentation.Text("Classic review")) {
                    AuthorId = 0U,
                    Index = 1U,
                    DateTime = new DateTime(2026, 7, 10, 8, 0, 0, DateTimeKind.Utc)
                });

            string modernAuthorId = "{11111111-1111-1111-1111-111111111111}";
            string modernCommentId = "{22222222-2222-2222-2222-222222222222}";
            string modernReplyId = "{33333333-3333-3333-3333-333333333333}";
            PowerPointAuthorsPart modernAuthors = presentationPart.AddNewPart<PowerPointAuthorsPart>();
            FeedXml(modernAuthors, $"""
                <p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">
                  <p188:author id="{modernAuthorId}" name="Modern Reviewer" initials="MR" userId="reviewer@example.test" providerId="OfficeIMO" />
                </p188:authorLst>
                """);
            PowerPointCommentPart modernPart = slide.SlidePart.AddNewPart<PowerPointCommentPart>();
            FeedXml(modernPart, $"""
                <p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p188:cm id="{modernCommentId}" authorId="{modernAuthorId}" status="active" created="2026-07-10T08:05:00Z">
                    <p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Modern review</a:t></a:r></a:p></p188:txBody>
                    <p188:replyLst>
                      <p188:reply id="{modernReplyId}" authorId="{modernAuthorId}" status="active" created="2026-07-10T08:06:00Z">
                        <p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Reply review</a:t></a:r></a:p></p188:txBody>
                      </p188:reply>
                    </p188:replyLst>
                  </p188:cm>
                </p188:cmLst>
                """);

            slide.SlidePart.Slide!.Timing = new Timing(
                new TimeNodeList(
                    new ParallelTimeNode(
                        new CommonTimeNode(
                            new ChildTimeNodeList(
                                new Animate(
                                    new CommonBehavior(
                                        new CommonTimeNode { Id = 2U, Duration = "500" },
                                        new TargetElement(new ShapeTarget {
                                            ShapeId = target.Id!.Value.ToString(CultureInfo.InvariantCulture)
                                        }))))) {
                            Id = 1U,
                            Duration = "indefinite",
                            NodeType = TimeNodeValues.TmingRoot
                        })));

            PowerPointReviewReport review = presentation.InspectReviewComments();
            PowerPointAnimationReport animation = presentation.InspectAnimations();

            Assert.Equal(1, review.ClassicCount);
            Assert.Equal(2, review.ModernCount);
            Assert.Contains(review.Comments, comment => comment.AuthorName == "Classic Reviewer" &&
                comment.Text == "Classic review");
            Assert.True(review.Comments.Any(comment => comment.AuthorName == "Modern Reviewer" &&
                comment.Text == "Modern review" && string.Equals(comment.Status, "Active",
                    StringComparison.OrdinalIgnoreCase)), review.ToJson());
            PowerPointReviewComment reply = Assert.Single(review.Comments,
                comment => comment.Kind == PowerPointCommentKind.ModernReply);
            Assert.Equal(modernCommentId, reply.ParentId);
            Assert.Equal("Reply review", reply.Text);
            Assert.Contains("\"commentCount\":3", review.ToJson(), StringComparison.Ordinal);
            PowerPointAnimationNode animated = Assert.Single(animation.Nodes,
                node => node.Kind == PowerPointAnimationKind.Animate);
            PowerPointAnimationNode container = Assert.Single(animation.Nodes,
                node => node.Kind == PowerPointAnimationKind.Parallel);
            Assert.Equal("1", container.TimingId);
            Assert.Null(container.ShapeId);
            Assert.Null(container.ShapeName);
            Assert.Equal(target.Id, animated.ShapeId);
            Assert.Equal(target.Name, animated.ShapeName);
            Assert.Equal("500", animated.Duration);
        }

        [Fact]
        public void AnimationInspectionBoundsTraversalAndProjectedNodes() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SlidePart.Slide!.Timing = new Timing(
                new TimeNodeList(
                    new ParallelTimeNode(new CommonTimeNode { Id = 1U }),
                    new ParallelTimeNode(new CommonTimeNode { Id = 2U })));

            InvalidDataException nodeException = Assert.Throws<InvalidDataException>(() =>
                presentation.InspectAnimations(new PowerPointAnimationInspectionOptions { MaxAnimationNodes = 1 }));
            InvalidDataException elementException = Assert.Throws<InvalidDataException>(() =>
                presentation.InspectAnimations(new PowerPointAnimationInspectionOptions { MaxXmlElements = 1 }));

            Assert.Contains(nameof(PowerPointAnimationInspectionOptions.MaxAnimationNodes), nodeException.Message, StringComparison.Ordinal);
            Assert.Contains(nameof(PowerPointAnimationInspectionOptions.MaxXmlElements), elementException.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SemanticSmartArtWorkflowsRoundTripEditableNodeText() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointSmartArt process = slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                        new[] { "Discover", "Design", "Deliver" }, 20, 20, 2600000, 1200000);
                    PowerPointSmartArt hierarchy = slide.AddSmartArt(PowerPointSmartArtType.BasicHierarchy,
                        new[] { "Executive", "Platform", "Delivery" }, 2800000, 20, 2600000, 1200000);
                    PowerPointSmartArt cycle = slide.AddSmartArt(PowerPointSmartArtType.BasicCycle,
                        new[] { "Plan", "Run", "Learn", "Improve" }, 20, 1400000, 5200000, 1400000);

                    Assert.Equal(new[] { "Discover", "Design", "Deliver" }, process.GetNodeTexts());
                    Assert.Equal(3, hierarchy.NodeCount);
                    Assert.Equal(4, cycle.NodeCount);
                    cycle.SetNodeText(3, "Adapt");
                    Assert.Equal("Adapt", cycle.GetNodeText(3));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointSmartArt[] diagrams = presentation.Slides[0].SmartArts.ToArray();
                    Assert.Equal(3, diagrams.Length);
                    Assert.Equal(new[] { "Discover", "Design", "Deliver" }, diagrams[0].GetNodeTexts());
                    Assert.Equal(new[] { "Plan", "Run", "Learn", "Adapt" }, diagrams[2].GetNodeTexts());
                    var validation = presentation.ValidateDocument();
                    Assert.True(validation.Count == 0, string.Join(Environment.NewLine,
                        validation.Select(error => error.Description + " | " + error.Path?.XPath)));
                }
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void NotesPagesAndHandoutsExportExistingNotesWithoutCreatingNewNotesParts() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            for (int index = 0; index < 3; index++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTitle("Slide " + (index + 1));
                if (index < 2) slide.Notes.Text = "Speaker note " + (index + 1);
            }
            Assert.Null(presentation.Slides[2].SlidePart.NotesSlidePart);

            byte[] notesPdf = presentation.ToPdf(new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.NotesPages,
                IncludeSpeakerNotes = true
            });
            byte[] handoutPdf = presentation.ToPdf(new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.Handouts,
                HandoutSlidesPerPage = 3,
                IncludeSpeakerNotes = true
            });

            using var notes = PdfPigDocument.Open(new MemoryStream(notesPdf));
            using var handout = PdfPigDocument.Open(new MemoryStream(handoutPdf));
            Assert.Equal(3, notes.NumberOfPages);
            Assert.Equal(1, handout.NumberOfPages);
            Assert.Contains("Speaker note 1", notes.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("Speaker note 2", handout.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Null(presentation.Slides[2].SlidePart.NotesSlidePart);
            Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointPdfSaveOptions {
                HandoutSlidesPerPage = 5
            });
        }

        [Theory]
        [InlineData(PowerPointPdfPageLayout.NotesPages)]
        [InlineData(PowerPointPdfPageLayout.Handouts)]
        public void NotesAndHandoutThumbnailsHonorPdfContentFilters(PowerPointPdfPageLayout layout) {
            using var controlStream = new MemoryStream();
            using var pictureStream = new MemoryStream();
            using PowerPointPresentation control = PowerPointPresentation.Create(controlStream, new PowerPointCreateOptions());
            using PowerPointPresentation withPicture = PowerPointPresentation.Create(pictureStream, new PowerPointCreateOptions());
            control.AddSlide().AddTitle("Filtered thumbnail");
            withPicture.AddSlide().AddTitle("Filtered thumbnail");
            withPicture.Slides[0].AddPicture(new MemoryStream(PdfPngTestImages.CreateRgbPng(255, 0, 0)),
                OfficeIMO.PowerPoint.ImagePartType.Png, PowerPointUnits.FromPoints(72), PowerPointUnits.FromPoints(72),
                PowerPointUnits.FromPoints(180), PowerPointUnits.FromPoints(120));

            var controlOptions = new PowerPointPdfSaveOptions { PageLayout = layout };
            var pictureOptions = new PowerPointPdfSaveOptions { PageLayout = layout };
            controlOptions.UseProfile(PdfCore.PdfExportProfile.TextOnly);
            pictureOptions.UseProfile(PdfCore.PdfExportProfile.TextOnly);

            byte[] controlThumbnail = PdfCore.PdfPageImageRenderer.RenderPageAsPng(
                control.ToPdf(controlOptions));
            byte[] pictureThumbnail = PdfCore.PdfPageImageRenderer.RenderPageAsPng(
                withPicture.ToPdf(pictureOptions));
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                controlThumbnail, pictureThumbnail, channelTolerance: 0, allowedDifferentPixels: 0);

            Assert.True(comparison.Passed,
                $"Filtered {layout} thumbnail changed at {comparison.DifferentPixels} pixels.");
        }

        [Fact]
        public void SignedPresentationSaveIsBlockedUntilMutationPolicyIsExplicit() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed workflow");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path)) {
                    PowerPointSignatureReport inspection = presentation.InspectSignatures();
                    Assert.True(inspection.HasSignatureMetadata);
                    Assert.Equal(1, inspection.XmlSignaturePartCount);
                    PowerPointSignedPresentationMutationException blocked =
                        Assert.Throws<PowerPointSignedPresentationMutationException>(() => presentation.Save());
                    Assert.Equal(PowerPointSignatureMutationAction.Blocked, blocked.Report.Action);

                    presentation.SignatureMutationPolicy =
                        PowerPointSignatureMutationPolicy.RemoveInvalidatedSignatures;
                    presentation.Save();
                    Assert.Equal(PowerPointSignatureMutationAction.Removed,
                        presentation.LastSignatureReport!.Action);
                }

                using PresentationDocument reopened = PresentationDocument.Open(path, false);
                Assert.Null(reopened.DigitalSignatureOriginPart);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SignedPresentationDisposeCannotBypassMutationPolicy() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed workflow");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                PowerPointPresentation edited = PowerPointPresentation.Load(path,
                    new PowerPointLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose });
                edited.Slides[0].AddTextBox("Must not persist");
                PowerPointSignedPresentationMutationException blocked =
                    Assert.Throws<PowerPointSignedPresentationMutationException>(() => edited.Dispose());

                Assert.Equal(PowerPointSignatureMutationAction.Blocked, blocked.Report.Action);
                using (PresentationDocument signed = PresentationDocument.Open(path, false)) {
                    Assert.NotNull(signed.DigitalSignatureOriginPart);
                }
                using PowerPointPresentation reopened = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
                Assert.DoesNotContain(reopened.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Must not persist");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void UntouchedSignedPresentationCanBeInspectedThroughEditableOpen() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed inspection");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    Assert.True(presentation.InspectSignatures().HasSignatureMetadata);
                    Assert.Equal("Signed inspection", presentation.Slides[0].TextBoxes.First().Text);
                }

                using PresentationDocument signed = PresentationDocument.Open(path, false);
                Assert.NotNull(signed.DigitalSignatureOriginPart);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SmartArtDeckCanUseOptInPowerPointDesktopReferenceLane() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_DESKTOP_REFERENCE"),
                    "1", StringComparison.Ordinal)) return;
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.SmartArtReference", Guid.NewGuid().ToString("N"));
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddSmartArt(PowerPointSmartArtType.BasicProcess,
                        new[] { "Inspect", "Build", "Validate" });
                    presentation.Save();
                }
                PowerPointReferenceRenderResult result = PowerPointDesktopReferenceRenderer.TryRender(path, output,
                    enabled: true);
                Assert.True(result.IsSuccessful, result.Message);
                Assert.NotEmpty(result.ImagePaths);
            } finally {
                if (File.Exists(path)) File.Delete(path);
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneRemovesOnlyStalePowerPointSlideImages() {
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.ReferenceCleanup",
                Guid.NewGuid().ToString("N"));
            try {
                Directory.CreateDirectory(output);
                File.WriteAllBytes(Path.Combine(output, "Slide1.png"), new byte[] { 1 });
                File.WriteAllBytes(Path.Combine(output, "slide12.PNG"), new byte[] { 2 });
                File.WriteAllBytes(Path.Combine(output, "comparison.png"), new byte[] { 3 });

                PowerPointDesktopReferenceRenderer.ClearExistingSlideImages(output);

                Assert.False(File.Exists(Path.Combine(output, "Slide1.png")));
                Assert.False(File.Exists(Path.Combine(output, "slide12.PNG")));
                Assert.True(File.Exists(Path.Combine(output, "comparison.png")));
            } finally {
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneReturnsOnlySlideImagesInNumericOrder() {
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.ReferenceOrder",
                Guid.NewGuid().ToString("N"));
            try {
                Directory.CreateDirectory(output);
                File.WriteAllBytes(Path.Combine(output, "Slide10.png"), new byte[] { 10 });
                File.WriteAllBytes(Path.Combine(output, "Slide2.PNG"), new byte[] { 2 });
                File.WriteAllBytes(Path.Combine(output, "Slide1.png"), new byte[] { 1 });
                File.WriteAllBytes(Path.Combine(output, "comparison.png"), new byte[] { 3 });

                string[] images = PowerPointDesktopReferenceRenderer.GetSlideImagesInOrder(output);

                Assert.Equal(new[] { "Slide1.png", "Slide2.PNG", "Slide10.png" },
                    images.Select(Path.GetFileName));
            } finally {
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        private static void FeedXml(OpenXmlPart part, string xml) {
            using var data = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(data);
        }

        private static void AddSyntheticSignature(string path) {
            using PresentationDocument document = PresentationDocument.Open(path, true);
            DigitalSignatureOriginPart origin = document.AddDigitalSignatureOriginPart();
            XmlSignaturePart signature = origin.AddNewPart<XmlSignaturePart>();
            FeedXml(signature,
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo/><SignatureValue>AA==</SignatureValue></Signature>");
        }
    }
}
