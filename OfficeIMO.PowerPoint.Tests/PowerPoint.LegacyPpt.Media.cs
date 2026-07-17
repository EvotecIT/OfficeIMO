using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void FreshEmbeddedWaveMedia_WritesNativeObjectAndProjectsExactly() {
            byte[] wave = CreateMediaWavePayload();
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var audio = new MemoryStream(wave,
                    writable: false);
                PowerPointMedia media = slide.AddAudio(audio, "audio/wav",
                    ".wav", 15875L, 31750L, 158750L, 79375L);
                media.Name = "Native WAV";
                slide.Transition = SlideTransition.Fade;
                using var transitionSound = new MemoryStream(wave,
                    writable: false);
                slide.SetTransitionSound(transitionSound, "Native WAV");

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(binary);
            LegacyPptMedia mediaObject = Assert.Single(neutral.Media);
            LegacyPptSound sound = Assert.Single(neutral.Sounds);
            Assert.Equal(LegacyPptMediaKind.EmbeddedWaveAudio,
                mediaObject.Kind);
            Assert.Equal(sound.Id, mediaObject.SoundId);
            Assert.Equal(wave, mediaObject.GetData());
            Assert.Equal(1, mediaObject.DurationMilliseconds);
            Assert.Equal(sound.Id, Assert.IsType<LegacyPptTransition>(
                neutral.Slides[0].Transition).SoundId);
            LegacyPptShape shape = Assert.Single(neutral.Slides[0].Shapes);
            Assert.Equal(LegacyPptShapeKind.Media, shape.Kind);
            Assert.Same(mediaObject, shape.Media);
            Assert.Equal(10, shape.Bounds.Left);
            Assert.Equal(20, shape.Bounds.Top);
            Assert.Equal(100, shape.Bounds.Width);
            Assert.Equal(50, shape.Bounds.Height);

            using var input = new MemoryStream(binary, writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            PowerPointMedia projectedMedia = Assert.IsType<PowerPointMedia>(
                Assert.Single(projected.Slides[0].Shapes));
            Assert.Equal(wave, projectedMedia.GetData());
            Assert.Equal(15875L, projectedMedia.Left);
            Assert.Equal(31750L, projectedMedia.Top);
            Assert.Equal(158750L, projectedMedia.Width);
            Assert.Equal(79375L, projectedMedia.Height);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(binary,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedEmbeddedWaveMedia_BlocksPayloadOnlyReplacement() {
            byte[] original = CreateMediaWavePayload();
            byte[] replacement = (byte[])original.Clone();
            replacement[replacement.Length - 1] ^= 0x20;
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var audio = new MemoryStream(original,
                    writable: false);
                slide.AddAudio(audio, "audio/wav", ".wav");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                PowerPointMedia media = Assert.Single(
                    imported.Slides[0].Media);
                using var updated = new MemoryStream(replacement,
                    writable: false);
                media.UpdateData(updated);

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.False(preflight.CanWrite);
                Assert.Contains(preflight.Findings, finding =>
                    finding.Code == "PPT-WRITE-IMPORT-LOSS");
                Assert.Throws<NotSupportedException>(() =>
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                Assert.Equal(replacement, media.GetData());
            }
        }

        [Theory]
        [InlineData(true, "video/mp4", ".mp4",
            LegacyPptFeature.EmbeddedVideo)]
        [InlineData(false, "audio/mpeg", ".mp3",
            LegacyPptFeature.Media)]
        public void NonRepresentableEmbeddedMedia_IsExplicitlyBlocked(
            bool video, string contentType, string extension,
            LegacyPptFeature expectedFeature) {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(
                P.SlideLayoutValues.Blank);
            using var data = new MemoryStream(new byte[] { 1, 2, 3, 4 },
                writable: false);
            if (video) {
                slide.AddVideo(data, contentType, extension);
            } else {
                slide.AddAudio(data, contentType, extension);
            }

            LegacyPptWritePreflightReport preflight = source
                .AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            LegacyPptWriteFinding finding = Assert.Single(
                preflight.Findings, item =>
                    item.Code == "PPT-WRITE-MEDIA");
            Assert.Equal(expectedFeature, finding.Feature);
            Assert.Throws<NotSupportedException>(() =>
                source.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NonDefaultEmbeddedMediaPlayback_IsExplicitlyBlocked() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(
                P.SlideLayoutValues.Blank);
            using var audio = new MemoryStream(CreateMediaWavePayload(),
                writable: false);
            slide.AddAudio(audio, "audio/wav", ".wav");
            Assert.Single(slide.SlidePart.Slide!
                .Descendants<P.CommonMediaNode>()).Volume = 70000;

            LegacyPptWritePreflightReport preflight = source
                .AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-MEDIA"
                && finding.Feature == LegacyPptFeature.Media);
        }

        [Fact]
        public void EmbeddedWaveMedia_ProjectsWithExactAudioAndRoundTripsUnchanged() {
            byte[] wave = CreateMediaWavePayload();
            byte[] embeddedOleBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                slide.Transition = SlideTransition.Fade;
                using (var audio = new MemoryStream(wave,
                           writable: false)) {
                    slide.SetTransitionSound(audio, "Media WAV");
                }
                using var storage = new MemoryStream(
                    CreateOleTestStorage("Media carrier"),
                    writable: false);
                slide.AddOleObject(storage, "Package", 12700L, 25400L,
                    2743200L, 1828800L);
                embeddedOleBytes = source.ToBytes(
                    PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation carrier = LegacyPptPresentation.Load(
                embeddedOleBytes);
            uint soundId = Assert.Single(carrier.Sounds).Id;
            byte[] sourceBytes = ConvertEmbeddedOleContainer(
                embeddedOleBytes,
                ExternalObjectFixtureKind.EmbeddedWaveMedia,
                linkedUpdateMode: 3, embeddedSoundId: soundId);

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptMedia media = Assert.Single(neutral.Media);
            Assert.Equal(LegacyPptMediaKind.EmbeddedWaveAudio,
                media.Kind);
            Assert.True(media.HasProjectableAudio);
            Assert.False(media.Loop);
            Assert.False(media.Rewind);
            Assert.False(media.Narration);
            Assert.Equal(soundId, media.SoundId);
            Assert.Equal(2500, media.DurationMilliseconds);
            Assert.Equal(wave, media.GetData());
            Assert.Same(Assert.Single(neutral.Sounds), media.Sound);
            Assert.Empty(neutral.OleObjects);
            LegacyPptShape shape = Assert.Single(neutral.Slides[0].Shapes);
            Assert.Equal(LegacyPptShapeKind.Media, shape.Kind);
            Assert.Same(media, shape.Media);
            Assert.DoesNotContain(neutral.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-MEDIA-",
                    StringComparison.Ordinal));

            LegacyPptImportReport report = neutral.CreateImportReport();
            Assert.Equal(1, report.MediaObjectCount);
            Assert.Equal(1, report.EmbeddedWaveMediaCount);
            Assert.Equal(1, report.ProjectableMediaCount);
            Assert.Equal(0, report.LinkedOrDeviceMediaCount);

            using var input = new MemoryStream(sourceBytes,
                writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            PowerPointMedia projectedMedia = Assert.IsType<PowerPointMedia>(
                Assert.Single(projected.Slides[0].Shapes));
            Assert.Equal(PowerPointMediaKind.Audio, projectedMedia.Kind);
            Assert.Equal("audio/wav", projectedMedia.MediaContentType);
            Assert.Equal(wave, projectedMedia.GetData());
            Assert.Equal(sourceBytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Empty(projected.ValidateDocument());
            PowerPointFeatureFinding finding = Assert.Single(projected
                .InspectFeatures().FindFeatures("Audio and video"));
            Assert.Equal(1, finding.Count);
        }

        [Fact]
        public void ImportedEmbeddedWaveMedia_GeometryAndUnrelatedTextEditUseIncrementalRecord() {
            byte[] wave = CreateMediaWavePayload();
            byte[] carrierBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                slide.AddTextBox("Editable companion");
                slide.Transition = SlideTransition.Cut;
                using (var audio = new MemoryStream(wave,
                           writable: false)) {
                    slide.SetTransitionSound(audio, "Preserved WAV");
                }
                using var storage = new MemoryStream(
                    CreateOleTestStorage("Preservation carrier"),
                    writable: false);
                slide.AddOleObject(storage, "Package", 15875L, 31750L,
                    158750L, 79375L);
                carrierBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            uint soundId = Assert.Single(LegacyPptPresentation.Load(
                carrierBytes).Sounds).Id;
            byte[] sourceBytes = ConvertEmbeddedOleContainer(carrierBytes,
                ExternalObjectFixtureKind.EmbeddedWaveMedia,
                linkedUpdateMode: 3, embeddedSoundId: soundId);
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptMedia originalMedia = Assert.Single(original.Media);
            LegacyPptShape originalMediaShape = Assert.Single(
                original.Slides[0].Shapes, shape => shape.Media != null);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                Assert.Equal(sourceBytes,
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                PowerPointTextBox text = Assert.Single(
                    imported.Slides[0].TextBoxes);
                PowerPointMedia media = Assert.Single(
                    imported.Slides[0].Media);
                text.Text = "Edited companion";
                media.Left += 15875L;

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptMedia savedMedia = Assert.Single(saved.Media);
            LegacyPptShape savedMediaShape = Assert.Single(
                saved.Slides[0].Shapes, shape => shape.Media != null);
            Assert.Equal(originalMedia.Id, savedMedia.Id);
            Assert.Equal(originalMedia.SoundId, savedMedia.SoundId);
            Assert.Equal(originalMedia.DurationMilliseconds,
                savedMedia.DurationMilliseconds);
            Assert.Equal(wave, savedMedia.GetData());
            Assert.Equal(originalMediaShape.Bounds.Left + 10,
                savedMediaShape.Bounds.Left);
            Assert.Contains(saved.Slides[0].Shapes,
                shape => shape.Text == "Edited companion");
            Assert.Equal(original.Package.PersistObjects[
                    original.Package.DocumentPersistId].RecordBytes,
                saved.Package.PersistObjects[
                    saved.Package.DocumentPersistId].RecordBytes);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            PowerPointMedia reopenedMedia = Assert.Single(
                reopened.Slides[0].Media);
            Assert.Equal(wave, reopenedMedia.GetData());
            Assert.Equal(31750L, reopenedMedia.Left);
            Assert.Equal("Edited companion", Assert.Single(
                reopened.Slides[0].TextBoxes).Text);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Theory]
        [InlineData(ExternalObjectFixtureKind.LinkedWaveMedia,
            LegacyPptMediaKind.LinkedWaveAudio, @"C:\Media\sample.wav")]
        [InlineData(ExternalObjectFixtureKind.MidiAudio,
            LegacyPptMediaKind.MidiAudio, @"C:\Media\sample.mid")]
        [InlineData(ExternalObjectFixtureKind.AviMovie,
            LegacyPptMediaKind.AviMovie, @"C:\Media\sample.avi")]
        [InlineData(ExternalObjectFixtureKind.MciMovie,
            LegacyPptMediaKind.MciMovie, @"C:\Media\sample.avi")]
        public void PathMedia_ImportsTypedAndReportsPreservation(
            ExternalObjectFixtureKind fixtureKind,
            LegacyPptMediaKind expectedKind, string expectedPath) {
            byte[] sourceBytes = CreateExternalObjectFixture(
                CreateOleTestStorage(expectedKind.ToString()), fixtureKind,
                compressed: false);

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptMedia media = Assert.Single(neutral.Media);
            Assert.Equal(expectedKind, media.Kind);
            Assert.Equal(expectedPath, media.Path);
            Assert.False(media.HasProjectableAudio);
            Assert.Contains(neutral.Slides[0].Shapes,
                shape => ReferenceEquals(shape.Media, media));
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-MEDIA-PRESERVED");
            Assert.Equal(1, neutral.CreateImportReport()
                .LinkedOrDeviceMediaCount);

            using var input = new MemoryStream(sourceBytes,
                writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.Empty(projected.Slides[0].Media);
            PowerPointFeatureFinding finding = Assert.Single(projected
                .InspectFeatures().FindFeatures("Legacy media metadata"));
            Assert.Equal(PowerPointFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Equal(1, finding.Count);
            Assert.Equal(sourceBytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void CdAudio_ImportsTypedPackedRange() {
            byte[] sourceBytes = CreateExternalObjectFixture(
                CreateOleTestStorage("CD audio"),
                ExternalObjectFixtureKind.CdAudio, compressed: false);

            LegacyPptMedia media = Assert.Single(
                LegacyPptPresentation.Load(sourceBytes).Media);
            Assert.Equal(LegacyPptMediaKind.CdAudio, media.Kind);
            Assert.Equal(0x01020304U, media.CdStart);
            Assert.Equal(0x05060708U, media.CdEnd);
            Assert.Null(media.Path);
        }

        [Fact]
        public void EmbeddedWavePlaybackFlags_AreTypedAndReportedPreserveOnly() {
            byte[] wave = CreateMediaWavePayload();
            byte[] carrierBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                slide.Transition = SlideTransition.Cut;
                using (var audio = new MemoryStream(wave,
                           writable: false)) {
                    slide.SetTransitionSound(audio, "Flags WAV");
                }
                using var storage = new MemoryStream(
                    CreateOleTestStorage("Flags carrier"),
                    writable: false);
                slide.AddOleObject(storage, "Package");
                carrierBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            uint soundId = Assert.Single(LegacyPptPresentation.Load(
                carrierBytes).Sounds).Id;
            byte[] sourceBytes = ConvertEmbeddedOleContainer(carrierBytes,
                ExternalObjectFixtureKind.EmbeddedWaveMedia,
                linkedUpdateMode: 3, embeddedSoundId: soundId,
                mediaFlags: 0x0007);

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptMedia media = Assert.Single(neutral.Media);
            Assert.True(media.Loop);
            Assert.True(media.Rewind);
            Assert.True(media.Narration);
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-MEDIA-PRESERVED");

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.Single(projected.InspectFeatures().FindFeatures(
                "Legacy media metadata"));
            Assert.Equal(sourceBytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void CapabilityContract_SeparatesEmbeddedAndLinkedMedia() {
            LegacyPptCapability embedded = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.Media);
            LegacyPptCapability linked = LegacyPptCapabilityCatalog.Get(
                LegacyPptFeature.LinkedMedia);

            Assert.Equal(LegacyPptCapabilityState.Native,
                embedded.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Native,
                embedded.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                embedded.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Native,
                embedded.PptxToBinary);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                linked.ImportToEditableModel);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                linked.NewBinaryWrite);
            Assert.Equal(LegacyPptCapabilityState.Preserved,
                linked.BinaryRoundTrip);
            Assert.Equal(LegacyPptCapabilityState.Blocked,
                linked.PptxToBinary);
        }

        private static byte[] CreateMediaWavePayload() => new byte[] {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F',
            40, 0, 0, 0,
            (byte)'W', (byte)'A', (byte)'V', (byte)'E',
            (byte)'f', (byte)'m', (byte)'t', (byte)' ',
            16, 0, 0, 0,
            1, 0, 1, 0,
            0x40, 0x1F, 0, 0,
            0x40, 0x1F, 0, 0,
            1, 0, 8, 0,
            (byte)'d', (byte)'a', (byte)'t', (byte)'a',
            4, 0, 0, 0,
            0x80, 0x90, 0x70, 0x80
        };
    }
}
