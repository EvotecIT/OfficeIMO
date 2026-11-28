using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using System.Text;
using System.Xml.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.Tests {
    public class PowerPointMediaTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CanAddEmbeddedVideoAndReloadAsMediaShape(bool sdkAlreadyOpen) {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    if (sdkAlreadyOpen) _ = presentation.OpenXmlDocument;
                    using MemoryStream videoStream = new(new byte[] { 0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50 });

                    PowerPointMedia media = slide.AddVideo(videoStream, "video/mp4", ".mp4");

                    Assert.Equal(PowerPointMediaKind.Video, media.Kind);
                    Assert.Equal(PowerPointShapeContentType.Media, media.ShapeContentType);
                    Assert.Equal("video/mp4", media.MediaContentType);
                    Assert.NotNull(media.MediaReferenceId);
                    Assert.NotNull(media.PlaybackReferenceId);
                    Assert.Equal(videoStream.ToArray(), media.GetData());
                    presentation.Save();
                    Assert.Equal(sdkAlreadyOpen, presentation.IsSdkLoaded);
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

                using (PowerPointDocument reloaded = PowerPointDocument.Load(filePath)) {
                    PowerPointMedia media = Assert.IsType<PowerPointMedia>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal(PowerPointMediaKind.Video, media.Kind);
                    Assert.Equal("video/mp4", media.MediaContentType);
                    Assert.False(reloaded.IsSdkLoaded);
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
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using MemoryStream audioStream = new(new byte[] { 73, 68, 51, 4, 0, 0, 0, 0, 0, 0 });

                    PowerPointMedia media = slide.AddAudio(audioStream, "audio/mpeg", ".mp3");

                    Assert.Equal(PowerPointMediaKind.Audio, media.Kind);
                    Assert.Equal(PowerPointShapeContentType.Media, media.ShapeContentType);
                    Assert.Equal("audio/mpeg", media.MediaContentType);
                    Assert.NotNull(media.MediaReferenceId);
                    Assert.NotNull(media.PlaybackReferenceId);
                    Assert.Equal(audioStream.ToArray(), media.GetData());
                    presentation.Save();
                    Assert.False(presentation.IsSdkLoaded);
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

                using (PowerPointDocument reloaded = PowerPointDocument.Load(filePath)) {
                    PowerPointMedia media = Assert.IsType<PowerPointMedia>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal(PowerPointMediaKind.Audio, media.Kind);
                    Assert.Equal("audio/mpeg", media.MediaContentType);
                    Assert.False(reloaded.IsSdkLoaded);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void LinkedMediaIgnoresPlaybackIdsInsideUnknownExtensionMarkup() {
            using var presentation = PowerPointDocument.Create();
            var source = new Uri("https://example.test/media%20clip.mp4?download=1");
            PowerPointMedia video = presentation.AddSlide().AddLinkedVideo(source);
            string playbackId = video.PlaybackReferenceId!;
            Assert.Equal(video.MediaReferenceId, playbackId);
            Assert.Equal(source, video.LinkUri);
            Assert.Equal(PowerPointMediaSourceKind.Linked, video.SourceKind);
            using var input = new MemoryStream(presentation.ToBytes());
            Assert.False(presentation.IsSdkLoaded);

            const string path = "/ppt/slides/slide1.xml";
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace future = "urn:test:future-media";
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(path)));
            XElement extensions = xml.Descendants(p + "nvPicPr").Single().Element(p + "nvPr")!.Element(p + "extLst")!;
            XAttribute mediaExtension = extensions.Elements(p + "ext").Single().Attribute("uri")!;
            string mediaUri = mediaExtension.Value;
            mediaExtension.Value = mediaUri.ToLowerInvariant();
            extensions.AddFirst(new XElement(p + "ext", new XAttribute("uri", mediaUri),
                new XElement(future + "producer", new XElement(p14 + "media", new XAttribute(r + "link", "rIdFake")))));
            byte[] modified = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(path, modified, modified.Length);
            using var output = new MemoryStream();
            package.Save(output);
            output.Position = 0;

            using PowerPointDocument reloaded = PowerPointDocument.Load(output);
            PowerPointMedia restored = Assert.Single(reloaded.Slides.Single().Media);
            Assert.Equal(playbackId, restored.PlaybackReferenceId);
            Assert.Equal(source, restored.LinkUri);
            Assert.Throws<InvalidOperationException>(() => restored.GetData());
            Assert.False(reloaded.IsSdkLoaded);
            _ = reloaded.OpenXmlDocument;
            Assert.Equal(playbackId, restored.PlaybackReferenceId);
        }

        [Fact]
        public void UnreadableVideoPosterDoesNotLeaveMediaPartsOrTiming() {
            using var presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var video = new MemoryStream(new byte[] { 0, 0, 0, 24, 102, 116, 121, 112 });
            var poster = new MemoryStream();
            poster.Dispose();

            Assert.Throws<ArgumentException>(() => slide.AddVideo(video, "video/mp4", ".mp4", posterImage: poster));
            Assert.Empty(slide.Media);
            Assert.False(presentation.IsSdkLoaded);
            using var bytes = new MemoryStream(presentation.ToBytes());
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(bytes);
            Assert.DoesNotContain(package.PartNames, name => name.Contains("/media/"));
            XDocument slideXml = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes("/ppt/slides/slide1.xml")));
            Assert.DoesNotContain(slideXml.Root!.Elements(), element => element.Name.LocalName == "timing");
        }

        [Fact]
        public void MediaPlaybackAndLinkEditsIgnoreTimingNodesInsideExtensions() {
            using var document = PowerPointDocument.Create();
            PowerPointMedia video = document.AddSlide().AddLinkedVideo(new Uri("https://example.test/original.mp4"));
            using var input = new MemoryStream(document.ToBytes());
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            const string slidePart = "/ppt/slides/slide1.xml";
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            XNamespace future = "urn:test:future-playback";
            XDocument markup = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(slidePart)));
            XElement media = markup.Descendants(p14 + "media").Single();
            var originalMedia = new XElement(media);
            originalMedia.Add(new XElement(p14 + "trim", new XAttribute("st", 75)));
            media.ReplaceWith(new XElement(mc + "AlternateContent",
                new XElement(mc + "Choice", new XAttribute(XNamespace.Xmlns + "future", future.NamespaceName),
                    new XAttribute("Requires", "future"),
                    new XElement(future + "producer", new XElement(p14 + "media", new XElement(p14 + "trim", new XAttribute("st", 10))))),
                new XElement(mc + "Fallback", originalMedia)));
            markup.Root!.Element(p + "timing")!.AddFirst(new XElement(p + "extLst",
                new XElement(p + "ext", new XAttribute("uri", "{0E537883-A8B0-4879-815E-773671E984D7}"),
                    new XElement(p + "cMediaNode", new XAttribute("vol", 8000), new XElement(p + "cTn", new XAttribute("id", 100)),
                        new XElement(p + "tgtEl", new XElement(p + "spTgt", new XAttribute("spid", video.Id!.Value)))))));
            markup.Root.Element(p + "timing")!.Element(p + "tnLst")!.AddFirst(new XElement(p + "video",
                new XElement(p + "cMediaNode", new XAttribute("vol", 9000), new XElement(p + "cTn", new XAttribute("id", 101)),
                    new XElement(p + "tgtEl", new XElement(p + "spTgt", new XAttribute("spid", video.Id!.Value))))));
            byte[] modified = Encoding.UTF8.GetBytes(markup.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(slidePart, modified, modified.Length);
            using var prepared = new MemoryStream();
            package.Save(prepared);
            prepared.Position = 0;

            using var loaded = PowerPointDocument.Load(prepared);
            PowerPointMedia frame = Assert.Single(loaded.Slides.Single().Media);
            Assert.Equal(80, frame.GetPlaybackOptions().VolumePercent);
            Assert.Equal(75U, frame.GetPlaybackOptions().TrimStartMilliseconds);
            var target = new Uri("https://example.test/updated%20video.mp4");
            frame.UpdateLink(target);
            frame.SetPlaybackOptions(new PowerPointMediaPlaybackOptions {
                VolumePercent = 35, Mute = true, Loop = true, FullScreen = true, ShowWhenStopped = false,
                TrimStartMilliseconds = 400, TrimEndMilliseconds = 2600, FadeInMilliseconds = 120, FadeOutMilliseconds = 80
            });
            using var saved = new MemoryStream(loaded.ToBytes());
            Assert.False(loaded.IsSdkLoaded);
            using OfficeOpcEditablePackage result = OfficeOpcEditablePackage.Open(saved, leaveOpen: true);
            XDocument output = XDocument.Parse(Encoding.UTF8.GetString(result.ReadPartBytes(slidePart)));
            XElement extension = output.Root!.Element(p + "timing")!.Element(p + "extLst")!;
            Assert.Equal("8000", (string?)extension.Descendants(p + "cMediaNode").Single().Attribute("vol"));
            XElement strayVideo = output.Root.Element(p + "timing")!.Element(p + "tnLst")!.Element(p + "video")!;
            Assert.Equal("9000", (string?)strayVideo.Element(p + "cMediaNode")!.Attribute("vol"));

            saved.Position = 0;
            using var reloaded = PowerPointDocument.Load(saved);
            PowerPointMedia persisted = Assert.Single(reloaded.Slides.Single().Media);
            PowerPointMediaPlaybackOptions options = persisted.GetPlaybackOptions();
            Assert.Equal(target, persisted.LinkUri);
            Assert.Equal(35, options.VolumePercent);
            Assert.True(options.Mute && options.Loop && options.FullScreen);
            Assert.False(options.ShowWhenStopped);
            Assert.Equal(400U, options.TrimStartMilliseconds);
            Assert.Equal(2600U, options.TrimEndMilliseconds);
            Assert.Equal(120U, options.FadeInMilliseconds);
            Assert.Equal(80U, options.FadeOutMilliseconds);
            Assert.False(reloaded.IsSdkLoaded);
        }

        [Fact]
        public void UpdatingImplicitFileMediaLinkRetainsTheAbsoluteUriUntilReload() {
            using var presentation = PowerPointDocument.Create();
            PowerPointMedia video = presentation.AddSlide().AddLinkedVideo(new Uri("https://example.test/old.mp4"));
            var fileUri = new Uri(Path.Combine(Path.GetTempPath(), "local video.mp4"), UriKind.Absolute);

            video.UpdateLink(fileUri);
            Assert.Equal(fileUri, video.LinkUri);
            using var stream = new MemoryStream(presentation.ToBytes());
            using var reloaded = PowerPointDocument.Load(stream);
            PowerPointMedia restored = Assert.Single(reloaded.Slides.Single().Media);
            Uri? uri = restored.LinkUri;
            Assert.Equal(fileUri.OriginalString, uri?.OriginalString);
            _ = reloaded.OpenXmlDocument;
            Assert.Equal(uri, restored.LinkUri);
        }
    }
}
