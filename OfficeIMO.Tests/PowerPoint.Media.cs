using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.Tests {
    public class PowerPointMediaTests {
        [Fact]
        public void CanAddEmbeddedVideoAndReloadAsMediaShape() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using MemoryStream videoStream = new(new byte[] { 0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50 });

                    PowerPointMedia media = slide.AddVideo(videoStream, "video/mp4", ".mp4");

                    Assert.Equal(PowerPointMediaKind.Video, media.Kind);
                    Assert.Equal(PowerPointShapeContentType.Media, media.ShapeContentType);
                    Assert.Equal("video/mp4", media.MediaContentType);
                    Assert.NotNull(media.MediaReferenceId);
                    Assert.NotNull(media.PlaybackReferenceId);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Assert.Single(slidePart.DataPartReferenceRelationships.OfType<VideoReferenceRelationship>());
                    Assert.Single(slidePart.DataPartReferenceRelationships.OfType<MediaReferenceRelationship>());

                    Picture picture = slidePart.Slide.Descendants<Picture>().Single();
                    ApplicationNonVisualDrawingProperties appProperties =
                        picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;
                    Assert.NotNull(appProperties.GetFirstChild<A.VideoFromFile>());
                    Assert.NotNull(appProperties.Descendants<P14.Media>().SingleOrDefault());
                    Assert.NotNull(slidePart.Slide.Timing?.Descendants<Video>().SingleOrDefault());
                }

                using (PowerPointPresentation reloaded = PowerPointPresentation.Open(filePath)) {
                    PowerPointMedia media = Assert.IsType<PowerPointMedia>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal(PowerPointMediaKind.Video, media.Kind);
                    Assert.Equal("video/mp4", media.MediaContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanAddEmbeddedAudioAndReloadAsMediaShape() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using MemoryStream audioStream = new(new byte[] { 73, 68, 51, 4, 0, 0, 0, 0, 0, 0 });

                    PowerPointMedia media = slide.AddAudio(audioStream, "audio/mpeg", ".mp3");

                    Assert.Equal(PowerPointMediaKind.Audio, media.Kind);
                    Assert.Equal(PowerPointShapeContentType.Media, media.ShapeContentType);
                    Assert.Equal("audio/mpeg", media.MediaContentType);
                    Assert.NotNull(media.MediaReferenceId);
                    Assert.NotNull(media.PlaybackReferenceId);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Assert.Single(slidePart.DataPartReferenceRelationships.OfType<AudioReferenceRelationship>());
                    Assert.Single(slidePart.DataPartReferenceRelationships.OfType<MediaReferenceRelationship>());

                    Picture picture = slidePart.Slide.Descendants<Picture>().Single();
                    ApplicationNonVisualDrawingProperties appProperties =
                        picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;
                    Assert.NotNull(appProperties.GetFirstChild<A.AudioFromFile>());
                    Assert.NotNull(appProperties.Descendants<P14.Media>().SingleOrDefault());
                    Assert.NotNull(slidePart.Slide.Timing?.Descendants<Audio>().SingleOrDefault());
                }

                using (PowerPointPresentation reloaded = PowerPointPresentation.Open(filePath)) {
                    PowerPointMedia media = Assert.IsType<PowerPointMedia>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal(PowerPointMediaKind.Audio, media.Kind);
                    Assert.Equal("audio/mpeg", media.MediaContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
