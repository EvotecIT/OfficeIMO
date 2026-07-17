using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public sealed class PowerPointMutationLifecycleTests {
        [Fact]
        public void RemovingOnlyCustomShowSlideRemovesShowAndActions() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointAutoShape actionShape = source.AddRectangle(
                100000, 100000, 1000000, 500000);
            PowerPointSlide target = presentation.AddSlide();
            PresentationPart presentationPart = presentation
                .OpenXmlDocument.PresentationPart!;
            const uint customShowId = 17U;
            presentationPart.Presentation!.CustomShowList =
                new CustomShowList(new CustomShow(
                    new SlideList(new SlideListEntry {
                        Id = presentationPart.GetIdOfPart(target.SlidePart)
                    })) {
                    Id = customShowId,
                    Name = "Removed show"
                });
            NonVisualDrawingProperties properties = ((Shape)actionShape.Element)
                .NonVisualShapeProperties!.NonVisualDrawingProperties!;
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                actionShape.SetClickSound(sound, "Removed show sound");
            }
            A.HyperlinkOnClick action = properties
                .GetFirstChild<A.HyperlinkOnClick>()!;
            action.Action = "ppaction://customshow?id=17&return=true";

            presentation.RemoveSlide(1);

            Assert.Null(presentationPart.Presentation.CustomShowList);
            Assert.Empty(source.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(source.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void DuplicatingMutuallyLinkedSlideSharesExistingTargets() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide first = presentation.AddSlide();
            PowerPointTextRun firstRun = first.AddTextBox("Second")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide second = presentation.AddSlide();
            PowerPointTextRun secondRun = second.AddTextBox("First")
                .Paragraphs.Single().Runs.Single();
            firstRun.SetHyperlink(second);
            secondRun.SetHyperlink(first);

            PowerPointSlide duplicate = presentation.DuplicateSlide(0);

            Assert.Equal(3, presentation.Slides.Count);
            Assert.Equal(3, presentation.OpenXmlDocument.PresentationPart!
                .SlideParts.Count());
            A.HyperlinkOnClick duplicateLink = duplicate.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.True(duplicate.SlidePart.TryGetPartById(
                duplicateLink.Id!.Value!, out OpenXmlPart? duplicateTarget));
            Assert.Same(second.SlidePart, duplicateTarget);
            A.HyperlinkOnClick secondLink = second.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.True(second.SlidePart.TryGetPartById(
                secondLink.Id!.Value!, out OpenXmlPart? secondTarget));
            Assert.Same(first.SlidePart, secondTarget);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ImportingMutuallyLinkedSlidesImportsListedTargetsOnce() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointTextRun firstRun = first.AddTextBox("Second")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide second = source.AddSlide();
            PowerPointTextRun secondRun = second.AddTextBox("First")
                .Paragraphs.Single().Runs.Single();
            firstRun.SetHyperlink(second);
            secondRun.SetHyperlink(first);
            using PowerPointPresentation target =
                PowerPointPresentation.Create();

            PowerPointSlide importedFirst = target.ImportSlide(source, 0);

            Assert.Equal(2, target.Slides.Count);
            PowerPointSlide importedSecond = target.Slides[1];
            A.HyperlinkOnClick firstLink = importedFirst.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            A.HyperlinkOnClick secondLink = importedSecond.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.True(importedFirst.SlidePart.TryGetPartById(
                firstLink.Id!.Value!, out OpenXmlPart? firstTarget));
            Assert.Same(importedSecond.SlidePart, firstTarget);
            Assert.True(importedSecond.SlidePart.TryGetPartById(
                secondLink.Id!.Value!, out OpenXmlPart? secondTarget));
            Assert.Same(importedFirst.SlidePart, secondTarget);
            Assert.Empty(target.ValidateDocument());
        }

        [Fact]
        public void RemovingSlideCleansLayoutAndSoundedInboundLinks() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide(
                P.SlideLayoutValues.Title);
            PowerPointSlide target = presentation.AddSlide();
            PowerPointTextRun run = source.AddTextBox("Target")
                .Paragraphs.Single().Runs.Single();
            run.SetHyperlink(target);
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                run.SetClickSound(sound, "Removed target sound");
            }
            SlideLayoutPart layoutPart = source.SlidePart.SlideLayoutPart!;
            const string LayoutRelationshipId = "rIdSlideTarget";
            layoutPart.AddPart(target.SlidePart, LayoutRelationshipId);
            NonVisualDrawingProperties layoutProperties = layoutPart
                .SlideLayout!.Descendants<NonVisualDrawingProperties>()
                .First();
            layoutProperties.Append(new A.HyperlinkOnClick {
                Id = LayoutRelationshipId
            });

            presentation.RemoveSlide(1);

            Assert.Empty(source.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(layoutPart.SlideLayout
                .Descendants<A.HyperlinkOnClick>());
            Assert.DoesNotContain(layoutPart.Parts, pair =>
                pair.RelationshipId == LayoutRelationshipId);
            Assert.Empty(source.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void DuplicatedNotesBacklinkTargetsDuplicateSlide() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide();
            source.Notes.Text = "Duplicated notes";
            source.SlidePart.NotesSlidePart!.AddPart(source.SlidePart);

            PowerPointSlide duplicate = presentation.DuplicateSlide(0);

            Assert.Same(source.SlidePart, source.SlidePart.NotesSlidePart!
                .SlidePart);
            Assert.Same(duplicate.SlidePart,
                duplicate.SlidePart.NotesSlidePart!.SlidePart);
            presentation.RemoveSlide(0);
            Assert.Same(duplicate.SlidePart,
                duplicate.SlidePart.NotesSlidePart!.SlidePart);
            Assert.Equal("Duplicated notes", duplicate.Notes.Text);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ShapeRemovalCleansClassicAnimationAndSound() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                slide.SetClassicAnimationSound(shape, sound,
                    "Removed animation");
            }

            shape.Remove();

            Assert.Empty(slide.Shapes);
            Assert.Empty(slide.ClassicAnimations);
            Assert.Null(slide.SlidePart.Slide!.Timing);
            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingClassicAnimationRejectsForeignShapeWithSameId() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide first = presentation.AddSlide();
            PowerPointAutoShape local = first.AddRectangle(
                100000, 100000, 1000000, 500000);
            first.AddClassicAnimation(local,
                PowerPointClassicAnimationEffect.Fade);
            PowerPointSlide second = presentation.AddSlide();
            PowerPointAutoShape foreign = second.AddRectangle(
                100000, 100000, 1000000, 500000);
            Assert.Equal(local.Id, foreign.Id);

            Assert.Throws<ArgumentException>(() =>
                first.RemoveClassicAnimation(foreign));

            Assert.Single(first.ClassicAnimations);
            Assert.Empty(second.ClassicAnimations);
        }

        [Fact]
        public void ClearingClassicAnimationsPreservesMediaTiming() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var audio = new MemoryStream(CreateWave(), writable: false);
            slide.AddAudio(audio, "audio/wav", ".wav");
            string originalTiming = slide.SlidePart.Slide!.Timing!.OuterXml;

            slide.ClearClassicAnimations();

            Assert.Equal(originalTiming,
                slide.SlidePart.Slide.Timing!.OuterXml);
            Assert.Single(slide.SlidePart.Slide.Timing.Descendants<Audio>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClearingClassicAnimationsRemovesOnlyClassicMixedTiming() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);
            using var audio = new MemoryStream(CreateWave(), writable: false);
            slide.AddAudio(audio, "audio/wav", ".wav");

            slide.ClearClassicAnimations();

            Assert.Empty(slide.ClassicAnimations);
            Assert.Empty(slide.SlidePart.Slide!.Timing!
                .Descendants<AnimateEffect>());
            Assert.Single(slide.SlidePart.Slide.Timing
                .Descendants<Audio>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void RemovingOneClassicAnimationPreservesMixedMediaTiming() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            PowerPointAutoShape second = slide.AddRectangle(
                100000, 700000, 1000000, 500000);
            slide.AddClassicAnimation(first,
                PowerPointClassicAnimationEffect.Fade);
            slide.AddClassicAnimation(second,
                PowerPointClassicAnimationEffect.Wipe);
            using var audio = new MemoryStream(CreateWave(), writable: false);
            slide.AddAudio(audio, "audio/wav", ".wav");
            string mediaTiming = Assert.Single(slide.SlidePart.Slide!
                .Timing!.Descendants<Audio>()).OuterXml;

            Assert.True(slide.RemoveClassicAnimation(first));

            PowerPointClassicAnimation remaining = Assert.Single(
                slide.ClassicAnimations);
            Assert.Equal(second.Id, remaining.ShapeId);
            Assert.Single(slide.SlidePart.Slide.Timing!
                .Descendants<AnimateEffect>());
            Assert.Equal(mediaTiming, Assert.Single(slide.SlidePart.Slide
                .Timing.Descendants<Audio>()).OuterXml);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ReplacingRunHyperlinkCleansOldTargetAndSoundRelationships() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointSlide target = presentation.AddSlide();
            PowerPointTextRun run = source.AddTextBox("Target")
                .Paragraphs.Single().Runs.Single();
            run.SetHyperlink("https://example.test/old");
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                run.SetClickSound(sound, "Old link sound");
            }
            Assert.Single(source.SlidePart.HyperlinkRelationships);
            Assert.Single(presentation.OpenXmlDocument.DataParts);

            run.SetHyperlink(target);

            Assert.Empty(source.SlidePart.HyperlinkRelationships);
            Assert.Empty(source.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Contains(source.SlidePart.Parts, pair => ReferenceEquals(
                pair.OpenXmlPart, target.SlidePart));

            run.ClearHyperlink();

            Assert.DoesNotContain(source.SlidePart.Parts, pair =>
                ReferenceEquals(pair.OpenXmlPart, target.SlidePart));
            Assert.Equal(2, presentation.Slides.Count);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void FailedSoundIngestionDoesNotLeaveMediaParts() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            PowerPointTextRun run = slide.AddTextBox("Sound")
                .Paragraphs.Single().Runs.Single();
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);

            Assert.Throws<IOException>(() => slide.SetTransitionSound(
                new ThrowingReadStream(CreateWave(), 8), "Broken"));
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Throws<IOException>(() => shape.SetClickSound(
                new ThrowingReadStream(CreateWave(), 8), "Broken"));
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Throws<IOException>(() => run.SetClickSound(
                new ThrowingReadStream(CreateWave(), 8), "Broken"));
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Throws<IOException>(() => slide.SetClassicAnimationSound(
                shape, new ThrowingReadStream(CreateWave(), 8), "Broken"));
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
        }

        [Fact]
        public void InvalidClassicAnimationSoundTargetDoesNotCreateMediaPart() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            using var sound = new MemoryStream(CreateWave(), writable: false);

            Assert.Throws<InvalidOperationException>(() =>
                slide.SetClassicAnimationSound(shape, sound, "No animation"));

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
        }

        [Fact]
        public void FailedMediaReplacementPreservesExistingPayload() {
            byte[] original = CreateWave();
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var initial = new MemoryStream(original, writable: false);
            PowerPointMedia media = slide.AddAudio(initial, "audio/wav",
                ".wav");

            Assert.Throws<IOException>(() => media.UpdateData(
                new ThrowingReadStream(original.Concat(
                    new byte[] { 1, 2, 3 }).ToArray(), 9)));

            Assert.Equal(original, media.GetData());
            Assert.Single(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        private static byte[] CreateWave() => new byte[] {
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

        private sealed class ThrowingReadStream : Stream {
            private readonly byte[] _bytes;
            private readonly int _failAfter;
            private int _position;

            internal ThrowingReadStream(byte[] bytes, int failAfter) {
                _bytes = bytes;
                _failAfter = failAfter;
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => _position;
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count) {
                if (_position >= _failAfter) {
                    throw new IOException("Injected read failure.");
                }
                int available = Math.Min(count,
                    Math.Min(_bytes.Length - _position,
                        _failAfter - _position));
                if (available <= 0) return 0;
                Buffer.BlockCopy(_bytes, _position, buffer, offset,
                    available);
                _position += available;
                return available;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) =>
                throw new NotSupportedException();
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset,
                int count) => throw new NotSupportedException();
        }
    }
}
