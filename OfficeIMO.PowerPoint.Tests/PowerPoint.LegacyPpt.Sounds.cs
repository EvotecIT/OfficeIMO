using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptSoundTests {
        [Fact]
        public void NativeWriter_DeduplicatesAndProjectsTransitionAndActionSound() {
            byte[] wave = CreateWavePayload();
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                slide.Transition = SlideTransition.Fade;
                MediaDataPart media = source.OpenXmlDocument.CreateMediaDataPart(
                    "audio/wav", ".wav");
                using (var input = new MemoryStream(wave, writable: false)) {
                    media.FeedData(input);
                }
                AudioReferenceRelationship soundRelationship = slide.SlidePart
                    .AddAudioReferenceRelationship(media);
                slide.SlidePart.Slide!.Transition!.Append(
                    new P.SoundAction(
                        new P.StartSoundAction(
                            new P.Sound {
                                Embed = soundRelationship.Id,
                                Name = "OfficeIMO Chime"
                            }) { Loop = true }));

                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship hyperlinkRelationship = slide.SlidePart
                    .AddHyperlinkRelationship(new Uri("https://example.com/sound"), true);
                P.NonVisualDrawingProperties properties = ((P.Shape)shape.Element)
                    .NonVisualShapeProperties!.NonVisualDrawingProperties!;
                properties.Append(new A.HyperlinkOnClick(
                    new A.HyperlinkSound {
                        Embed = soundRelationship.Id,
                        Name = "OfficeIMO Chime"
                    }) { Id = hyperlinkRelationship.Id });

                LegacyPptWritePreflightReport preflight = source.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptSound sound = Assert.Single(legacy.Sounds);
            Assert.Equal(1U, sound.Id);
            Assert.Equal("OfficeIMO Chime", sound.Name);
            Assert.Equal(".wav", sound.Extension);
            Assert.Equal("audio/wav", sound.ContentType);
            Assert.Equal(wave, sound.DataBytes);
            LegacyPptSlide binarySlide = Assert.Single(legacy.Slides);
            LegacyPptTransition transition = Assert.IsType<LegacyPptTransition>(
                binarySlide.Transition);
            Assert.True(transition.PlaySound);
            Assert.True(transition.LoopSound);
            Assert.False(transition.StopSound);
            Assert.Equal(sound.Id, transition.SoundId);
            LegacyPptInteraction interaction = Assert.Single(binarySlide.Shapes
                .SelectMany(item => item.Interactions));
            Assert.Equal(sound.Id, interaction.SoundIdReference);
            Assert.Equal(1, legacy.CreateImportReport().SoundCount);
            Assert.Equal(1, legacy.CreateImportReport().ImportableSoundCount);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-SOUND-", StringComparison.Ordinal));

            using var inputStream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(inputStream);
            PowerPointSlide projectedSlide = Assert.Single(projected.Slides);
            P.StartSoundAction projectedTransitionSound = Assert.Single(
                projectedSlide.SlidePart.Slide!.Transition!
                    .Descendants<P.StartSoundAction>());
            Assert.True(projectedTransitionSound.Loop?.Value);
            P.Sound projectedSound = Assert.Single(
                projectedTransitionSound.Elements<P.Sound>());
            A.HyperlinkSound projectedActionSound = Assert.Single(
                projectedSlide.SlidePart.Slide.Descendants<A.HyperlinkSound>());
            Assert.Equal(projectedSound.Embed?.Value,
                projectedActionSound.Embed?.Value);
            AudioReferenceRelationship projectedRelationship = Assert.Single(
                projectedSlide.SlidePart.DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>());
            Assert.Equal(projectedRelationship.Id, projectedSound.Embed?.Value);
            using Stream projectedBytes = projectedRelationship.DataPart.GetStream(
                FileMode.Open, FileAccess.Read);
            using var copied = new MemoryStream();
            projectedBytes.CopyTo(copied);
            Assert.Equal(wave, copied.ToArray());
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_AuthorsAndProjectsStopSoundTransition() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                slide.Transition = SlideTransition.Cut;
                slide.SlidePart.Slide!.Transition!.Append(
                    new P.SoundAction(new P.EndSoundAction()));
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptTransition transition = Assert.IsType<LegacyPptTransition>(
                Assert.Single(legacy.Slides).Transition);
            Assert.True(transition.StopSound);
            Assert.False(transition.PlaySound);
            Assert.Equal(0U, transition.SoundId);
            Assert.Empty(legacy.Sounds);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            Assert.Single(projected.Slides[0].SlidePart.Slide!.Transition!
                .Descendants<P.EndSoundAction>());
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void ImportedSoundPresentation_UnrelatedEditPreservesSoundSemantics() {
            byte[] sourceBytes = CreateBinarySoundPresentation();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                imported.Slides[0].Shapes[0].Left += 1588;
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptSound sound = Assert.Single(saved.Sounds);
            LegacyPptSlide slide = Assert.Single(saved.Slides);
            Assert.Equal(sound.Id, Assert.IsType<LegacyPptTransition>(
                slide.Transition).SoundId);
            Assert.Equal(sound.Id, Assert.Single(slide.Shapes
                .SelectMany(shape => shape.Interactions)).SoundIdReference);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedTransitionSoundEdit_AppendsSoundAndPatchesReference() {
            byte[] sourceBytes = CreateBinarySoundPresentation();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            byte[] replacement = CreateWavePayload().Concat(new byte[] {
                0x81, 0x82
            }).ToArray();

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                MediaDataPart media = imported.OpenXmlDocument.CreateMediaDataPart(
                    "audio/wav", ".wav");
                using (var audio = new MemoryStream(replacement, writable: false)) {
                    media.FeedData(audio);
                }
                AudioReferenceRelationship relationship = slide.SlidePart
                    .AddAudioReferenceRelationship(media);
                P.Sound sound = Assert.Single(slide.SlidePart.Slide!.Transition!
                    .Descendants<P.Sound>());
                sound.Embed = relationship.Id;
                sound.Name = "Replacement Sound";

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.True(saved.Sounds.Count == 2,
                string.Join(" | ", saved.Sounds.Select(sound =>
                    $"{sound.Id}:{sound.Name}:{sound.DataBytes.Length}")));
            LegacyPptSound appended = saved.Sounds.Single(sound =>
                sound.Name == "Replacement Sound");
            Assert.Equal(replacement, appended.DataBytes);
            LegacyPptTransition transition = Assert.IsType<LegacyPptTransition>(
                Assert.Single(saved.Slides).Transition);
            Assert.Equal(appended.Id, transition.SoundId);
            Assert.True(transition.PlaySound);
            Assert.Equal(saved.Sounds.Single(sound => sound.Name == "OfficeIMO Chime").Id,
                Assert.Single(saved.Slides[0].Shapes.SelectMany(shape =>
                    shape.Interactions)).SoundIdReference);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedShapeActionSoundEdit_AppendsSoundAndPatchesReference() {
            byte[] sourceBytes = CreateBinarySoundPresentation();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            byte[] replacement = CreateWavePayload().Concat(new byte[] {
                0x91, 0x92, 0x93
            }).ToArray();

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointShape shape = imported.Slides[0].Shapes.Single(item =>
                    item.ClickSoundName == "OfficeIMO Chime");
                using var audio = new MemoryStream(replacement, writable: false);
                shape.SetClickSound(audio, "Replacement Click");

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal(2, saved.Sounds.Count);
            LegacyPptSound originalSound = saved.Sounds.Single(sound =>
                sound.Name == "OfficeIMO Chime");
            LegacyPptSound appended = saved.Sounds.Single(sound =>
                sound.Name == "Replacement Click");
            Assert.Equal(replacement, appended.DataBytes);
            LegacyPptSlide slide = Assert.Single(saved.Slides);
            Assert.Equal(originalSound.Id, Assert.IsType<LegacyPptTransition>(
                slide.Transition).SoundId);
            Assert.Equal(appended.Id, Assert.Single(slide.Shapes
                .SelectMany(shape => shape.Interactions)).SoundIdReference);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ShapeActionSound_ReplacementAndClearReleaseMediaParts() {
            byte[] first = CreateWavePayload();
            byte[] second = first.Concat(new byte[] { 0x31, 0x32 })
                .ToArray();
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000L, 100000L, 800000L, 500000L);
            using (var audio = new MemoryStream(first, writable: false)) {
                shape.SetClickSound(audio, "First click");
            }
            Assert.Single(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Single(presentation.OpenXmlDocument.DataParts);

            using (var audio = new MemoryStream(second, writable: false)) {
                shape.SetClickSound(audio, "Second click");
            }

            Assert.Equal("Second click", shape.ClickSoundName);
            Assert.Equal(second, shape.GetClickSoundBytes());
            Assert.Single(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Single(presentation.OpenXmlDocument.DataParts);

            shape.ClearClickSound();

            Assert.Null(shape.ClickSoundName);
            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingTransitionReleasesItsMediaPart() {
            byte[] wave = CreateWavePayload();
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Transition = SlideTransition.Fade;
            using (var audio = new MemoryStream(wave, writable: false)) {
                slide.SetTransitionSound(audio, "Removed transition");
            }
            Assert.Single(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Single(presentation.OpenXmlDocument.DataParts);

            slide.Transition = SlideTransition.None;

            Assert.False(slide.HasTransitionSound);
            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingTransitionReleasesSoundsFromAllFallbackBranches() {
            byte[] wave = CreateWavePayload();
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Transition = SlideTransition.Morph;
            using (var audio = new MemoryStream(wave, writable: false)) {
                slide.SetTransitionSound(audio, "Choice sound");
            }
            MediaDataPart fallbackMedia = presentation.OpenXmlDocument
                .CreateMediaDataPart("audio/wav", ".wav");
            using (var audio = new MemoryStream(
                       wave.Concat(new byte[] { 0x21 }).ToArray(),
                       writable: false)) {
                fallbackMedia.FeedData(audio);
            }
            AudioReferenceRelationship fallbackRelationship = slide
                .SlidePart.AddAudioReferenceRelationship(fallbackMedia);
            AlternateContent alternate = Assert.Single(slide.SlidePart
                .Slide!.Elements<AlternateContent>());
            P.Transition fallbackTransition = alternate
                .GetFirstChild<AlternateContentFallback>()!
                .GetFirstChild<P.Transition>()!;
            fallbackTransition.GetFirstChild<P.SoundAction>()!
                .GetFirstChild<P.StartSoundAction>()!.Sound!.Embed =
                fallbackRelationship.Id;
            Assert.Equal(2, slide.SlidePart
                .DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>().Count());
            Assert.Equal(2, presentation.OpenXmlDocument.DataParts.Count());

            slide.Transition = SlideTransition.None;

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PublicApis_AuthorSoundOnlyTransitionShapeAndTextActions() {
            byte[] wave = CreateWavePayload();
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                using (var audio = new MemoryStream(wave, writable: false)) {
                    slide.SetTransitionSound(audio, "Shared Sound", loop: true);
                }
                slide.Transition = SlideTransition.WheelFourSpokes;
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                using (var audio = new MemoryStream(wave, writable: false)) {
                    shape.SetClickSound(audio, "Shared Sound");
                }
                PowerPointTextBox textBox = slide.AddTextBox("Sound text",
                    100000, 800000, 1600000, 500000);
                PowerPointTextRun run = textBox.Paragraphs.Single().Runs.Single();
                using (var audio = new MemoryStream(wave, writable: false)) {
                    run.SetMouseOverSound(audio, "Shared Sound");
                }
                run.SetMouseOverStopsSound(true);

                Assert.Equal("Shared Sound", slide.TransitionSoundName);
                Assert.True(slide.TransitionSoundLoops);
                Assert.Equal(wave, slide.GetTransitionSoundBytes());
                Assert.Equal("Shared Sound", shape.ClickSoundName);
                Assert.Equal(wave, shape.GetClickSoundBytes());
                Assert.Equal("Shared Sound", run.MouseOverSoundName);
                Assert.Equal(wave, run.GetMouseOverSoundBytes());
                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptSound sound = Assert.Single(legacy.Sounds);
            LegacyPptSlide binarySlide = Assert.Single(legacy.Slides);
            Assert.Equal(sound.Id, Assert.IsType<LegacyPptTransition>(
                binarySlide.Transition).SoundId);
            LegacyPptInteraction[] interactions = binarySlide.Shapes
                .SelectMany(shape => shape.Interactions.Concat(
                    shape.TextBody.Interactions.Select(item => item.Interaction)))
                .ToArray();
            Assert.Equal(2, interactions.Length);
            Assert.All(interactions, interaction => {
                Assert.Equal(LegacyPptInteractionAction.None, interaction.Action);
                Assert.Equal(sound.Id, interaction.SoundIdReference);
            });
            Assert.Contains(interactions, interaction => interaction.StopsSound);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            PowerPointSlide projectedSlide = Assert.Single(projected.Slides);
            Assert.Equal("Shared Sound", projectedSlide.TransitionSoundName);
            Assert.True(projectedSlide.TransitionSoundLoops);
            Assert.Equal(wave, projectedSlide.GetTransitionSoundBytes());
            Assert.Contains(projectedSlide.Shapes,
                item => item.ClickSoundName == "Shared Sound");
            PowerPointTextRun projectedRun = projectedSlide.Shapes
                .OfType<PowerPointTextBox>().Single().Paragraphs.Single()
                .Runs.Single();
            Assert.Equal("Shared Sound", projectedRun.MouseOverSoundName);
            Assert.Equal(wave, projectedRun.GetMouseOverSoundBytes());
            Assert.Empty(projected.ValidateDocument());
        }

        [Theory]
        [InlineData("audio/mpeg", ".mp3", false, "WAV or AIFF")]
        [InlineData("audio/wav", ".wav", true, "numeric")]
        public void PptxConversion_BlocksUnrepresentableTransitionSound(
            string contentType, string extension, bool builtIn,
            string expectedReason) {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            slide.Transition = SlideTransition.Fade;
            MediaDataPart media = source.OpenXmlDocument.CreateMediaDataPart(
                contentType, extension);
            using (var audio = new MemoryStream(CreateWavePayload(),
                       writable: false)) {
                media.FeedData(audio);
            }
            AudioReferenceRelationship relationship = slide.SlidePart
                .AddAudioReferenceRelationship(media);
            slide.SlidePart.Slide!.Transition!.Append(new P.SoundAction(
                new P.StartSoundAction(new P.Sound {
                    Embed = relationship.Id,
                    Name = "Unsupported Sound",
                    BuiltIn = builtIn
                })));

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-TRANSITION");
            Assert.Contains(expectedReason, finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] CreateBinarySoundPresentation() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            slide.Transition = SlideTransition.Fade;
            MediaDataPart media = source.OpenXmlDocument.CreateMediaDataPart(
                "audio/wav", ".wav");
            using (var input = new MemoryStream(CreateWavePayload(), writable: false)) {
                media.FeedData(input);
            }
            AudioReferenceRelationship soundRelationship = slide.SlidePart
                .AddAudioReferenceRelationship(media);
            slide.SlidePart.Slide!.Transition!.Append(new P.SoundAction(
                new P.StartSoundAction(new P.Sound {
                    Embed = soundRelationship.Id,
                    Name = "OfficeIMO Chime"
                }) { Loop = true }));
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            HyperlinkRelationship hyperlink = slide.SlidePart
                .AddHyperlinkRelationship(new Uri("https://example.com/sound"), true);
            ((P.Shape)shape.Element).NonVisualShapeProperties!
                .NonVisualDrawingProperties!.Append(new A.HyperlinkOnClick(
                    new A.HyperlinkSound {
                        Embed = soundRelationship.Id,
                        Name = "OfficeIMO Chime"
                    }) { Id = hyperlink.Id });
            return source.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static byte[] CreateWavePayload() => new byte[] {
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
