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
        public void ImportAndExportPreserveNotesAlongsideAudioRelationships() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide sourceSlide = source.AddSlide();
            using (var audio = new MemoryStream(CreateWave(),
                       writable: false)) {
                sourceSlide.AddAudio(audio, "audio/wav", ".wav");
            }
            sourceSlide.Notes.Text = "Media notes";
            using PowerPointPresentation target =
                PowerPointPresentation.Create();

            PowerPointSlide imported = target.ImportSlide(source, 0);

            Assert.Equal("Media notes", imported.Notes.Text);
            Assert.Single(imported.Media);
            Assert.Empty(target.ValidateDocument());

            using var exportedStream = new MemoryStream();
            source.ExportSlide(0, exportedStream);
            exportedStream.Position = 0;
            using PowerPointPresentation exported =
                PowerPointPresentation.Load(exportedStream);
            PowerPointSlide exportedSlide = Assert.Single(exported.Slides);
            Assert.Equal("Media notes", exportedSlide.Notes.Text);
            Assert.Single(exportedSlide.Media);
            Assert.Empty(exported.ValidateDocument());
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
            NotesSlidePart sourceNotesPart = source.SlidePart.NotesSlidePart;
            string backlinkId = sourceNotesPart.GetIdOfPart(
                source.SlidePart);
            sourceNotesPart.NotesSlide!
                .Descendants<NonVisualDrawingProperties>().First()
                .Name = backlinkId;

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
        public void ShapeRemovalCleansActionSoundMedia() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                shape.SetClickSound(sound, "Removed action sound");
            }

            shape.Remove();

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ShapeRemovalIgnoresOrdinaryAttributesMatchingLayoutId() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            SlideLayoutPart layout = slide.SlidePart.SlideLayoutPart!;
            string layoutRelationshipId = slide.SlidePart.GetIdOfPart(layout);
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            shape.Name = layoutRelationshipId;

            shape.Remove();

            Assert.Same(layout, slide.SlidePart.SlideLayoutPart);
            Assert.Contains(slide.SlidePart.Parts, pair =>
                pair.RelationshipId == layoutRelationshipId
                && ReferenceEquals(pair.OpenXmlPart, layout));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedMediaUpdateDetachesOnlyTheSelectedFrame() {
            byte[] original = CreateWave();
            byte[] replacement = (byte[])original.Clone();
            replacement[replacement.Length - 1] ^= 0x11;
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var input = new MemoryStream(original, writable: false);
            PowerPointMedia first = slide.AddAudio(input, "audio/wav",
                ".wav");
            first.Name = first.MediaReferenceId!;
            string originalName = first.Name!;
            PowerPointMedia second = Assert.IsType<PowerPointMedia>(
                first.Duplicate(offsetX: 1000000));
            Assert.Equal(first.MediaReferenceId,
                second.MediaReferenceId);

            using var updated = new MemoryStream(replacement,
                writable: false);
            first.UpdateData(updated);

            Assert.Equal(replacement, first.GetData());
            Assert.Equal(original, second.GetData());
            Assert.NotEqual(first.MediaReferenceId,
                second.MediaReferenceId);
            Assert.Equal(originalName, first.Name);
            Assert.Equal(2, presentation.OpenXmlDocument.DataParts.Count());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void MediaUpdateDetachesFromMasterConsumer() {
            byte[] original = CreateWave();
            byte[] replacement = (byte[])original.Clone();
            replacement[replacement.Length - 1] ^= 0x22;
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var input = new MemoryStream(original, writable: false);
            PowerPointMedia media = slide.AddAudio(input, "audio/wav",
                ".wav");
            MediaDataPart sharedPart = Assert.IsType<MediaDataPart>(slide
                .SlidePart.DataPartReferenceRelationships.Single(
                    relationship => relationship.Id ==
                        media.MediaReferenceId).DataPart);
            SlideMasterPart master = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!;
            AudioReferenceRelationship masterRelationship = master
                .AddAudioReferenceRelationship(sharedPart,
                    "rIdMasterMedia");
            master.SlideMaster!.CommonSlideData!.ShapeTree!.Append(
                CreateSoundedShape(202U, "Master media consumer",
                    masterRelationship.Id, "Shared master media"));
            master.SlideMaster.Save();

            using var updated = new MemoryStream(replacement,
                writable: false);
            media.UpdateData(updated);

            Assert.Equal(replacement, media.GetData());
            using Stream masterData = masterRelationship.DataPart.GetStream(
                FileMode.Open, FileAccess.Read);
            using var copied = new MemoryStream();
            masterData.CopyTo(copied);
            Assert.Equal(original, copied.ToArray());
            Assert.Equal(2, presentation.OpenXmlDocument.DataParts.Count());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ImportSlideClonesSharedMasterAndLayoutActionSound() {
            byte[] wave = CreateWave();
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide sourceSlide = source.AddSlide();
            SlideLayoutPart sourceLayout = sourceSlide.SlidePart
                .SlideLayoutPart!;
            sourceLayout.SlideLayout!.CommonSlideData!.Name =
                "Sounded import layout";
            SlideMasterPart sourceMaster = sourceLayout.SlideMasterPart!;
            MediaDataPart media = source.OpenXmlDocument
                .CreateMediaDataPart("audio/wav", ".wav");
            using (var input = new MemoryStream(wave, writable: false)) {
                media.FeedData(input);
            }
            AudioReferenceRelationship masterRelationship = sourceMaster
                .AddAudioReferenceRelationship(media, "rIdMasterSound");
            AudioReferenceRelationship layoutRelationship = sourceLayout
                .AddAudioReferenceRelationship(media, "rIdLayoutSound");
            sourceMaster.SlideMaster!.CommonSlideData!.ShapeTree!.Append(
                CreateSoundedShape(200U, "Master sound",
                    masterRelationship.Id, "Shared master sound"));
            sourceLayout.SlideLayout.CommonSlideData!.ShapeTree!.Append(
                CreateSoundedShape(201U, "Layout sound",
                    layoutRelationship.Id, "Shared layout sound"));
            sourceMaster.SlideMaster.Save();
            sourceLayout.SlideLayout.Save();
            using PowerPointPresentation target =
                PowerPointPresentation.Create();

            PowerPointSlide imported = target.ImportSlide(source, 0);

            SlideLayoutPart importedLayout = imported.SlidePart
                .SlideLayoutPart!;
            SlideMasterPart importedMaster = importedLayout
                .SlideMasterPart!;
            AudioReferenceRelationship importedMasterRelationship =
                Assert.Single(importedMaster.DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>());
            AudioReferenceRelationship importedLayoutRelationship =
                Assert.Single(importedLayout.DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>());
            Assert.Same(importedMasterRelationship.DataPart,
                importedLayoutRelationship.DataPart);
            Assert.Equal(importedMasterRelationship.Id, importedMaster
                .SlideMaster!.Descendants<A.HyperlinkSound>().Single()
                .Embed!.Value);
            Assert.Equal(importedLayoutRelationship.Id, importedLayout
                .SlideLayout!.Descendants<A.HyperlinkSound>().Single()
                .Embed!.Value);
            Assert.Single(target.OpenXmlDocument.DataParts);
            using Stream importedData = importedMasterRelationship.DataPart
                .GetStream(FileMode.Open, FileAccess.Read);
            using var copied = new MemoryStream();
            importedData.CopyTo(copied);
            Assert.Equal(wave, copied.ToArray());
            Assert.Empty(target.ValidateDocument());
        }

        [Fact]
        public void FailedLinkedSlideImportLeavesDestinationUnchanged() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointTextRun link = first.AddTextBox("Broken target")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide brokenTarget = source.AddSlide();
            link.SetHyperlink(brokenTarget);
            brokenTarget.SlidePart.DeletePart(
                brokenTarget.SlidePart.SlideLayoutPart!);
            using PowerPointPresentation target =
                PowerPointPresentation.Create();
            int originalPartCount = target.OpenXmlDocument
                .PresentationPart!.Parts.Count();
            string originalPresentationXml = target.OpenXmlDocument
                .PresentationPart.Presentation!.OuterXml;

            Assert.Throws<InvalidOperationException>(() =>
                target.ImportSlide(source, 0));

            Assert.Empty(target.Slides);
            Assert.Equal(originalPartCount, target.OpenXmlDocument
                .PresentationPart.Parts.Count());
            Assert.Equal(originalPresentationXml, target.OpenXmlDocument
                .PresentationPart.Presentation!.OuterXml);
            Assert.Empty(target.ValidateDocument());
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
        public void RemovingClassicAnimationPreservesMatchingAdvancedEffect() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);
            Timing timing = slide.SlidePart.Slide!.Timing!;
            AnimateEffect classicEffect = Assert.Single(
                timing.Descendants<AnimateEffect>());
            AnimateEffect advancedEffect = (AnimateEffect)
                classicEffect.CloneNode(true);
            foreach (CommonTimeNode timeNode in advancedEffect
                         .Descendants<CommonTimeNode>()) {
                if (timeNode.Id?.Value is uint id) timeNode.Id = id + 100U;
            }
            advancedEffect.CommonBehavior!.CommonTimeNode!.Duration = "777";
            var visibility = new SetBehavior(
                new CommonBehavior(
                    new CommonTimeNode {
                        Id = 250U,
                        Duration = "1",
                        Fill = TimeNodeFillValues.Hold
                    },
                    new TargetElement(new ShapeTarget {
                        ShapeId = shape.Id!.Value.ToString(
                            System.Globalization.CultureInfo.InvariantCulture)
                    }),
                    new AttributeNameList(new AttributeName(
                        "style.visibility"))),
                new ToVariantValue(new StringVariantValue {
                    Val = "visible"
                }));
            classicEffect.Parent!.InsertBefore(visibility, classicEffect);
            classicEffect.Parent.InsertBefore(
                (SetBehavior)visibility.CloneNode(true), classicEffect);
            classicEffect.Parent!.Append(advancedEffect);

            Assert.True(slide.RemoveClassicAnimation(shape));

            AnimateEffect remaining = Assert.Single(timing
                .Descendants<AnimateEffect>());
            Assert.Equal("777", remaining.CommonBehavior!
                .CommonTimeNode!.Duration!.Value);
            Assert.Empty(timing.Descendants<SetBehavior>());
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
        public void ClearingRunHyperlinkIgnoresShapeNameMatchingRelationshipId() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextRun run = slide.AddTextBox("Link")
                .Paragraphs.Single().Runs.Single();
            run.SetHyperlink("https://example.test/remove");
            string relationshipId = Assert.Single(slide.SlidePart
                .HyperlinkRelationships).Id;
            slide.AddRectangle(100000, 100000, 1000000, 500000)
                .Name = relationshipId;

            run.ClearHyperlink();

            Assert.Empty(slide.SlidePart.HyperlinkRelationships);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClearingShapeSoundIgnoresShapeNameMatchingRelationshipId() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape sounded = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                sounded.SetClickSound(sound, "Remove sound");
            }
            string relationshipId = Assert.Single(slide.SlidePart
                .DataPartReferenceRelationships).Id;
            slide.AddRectangle(100000, 700000, 1000000, 500000)
                .Name = relationshipId;

            sounded.ClearClickSound();

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
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

        private static Shape CreateSoundedShape(uint id, string name,
            string relationshipId, string soundName) {
            var properties = new NonVisualDrawingProperties {
                Id = id,
                Name = name
            };
            properties.Append(new A.HyperlinkOnClick(
                new A.HyperlinkSound {
                    Embed = relationshipId,
                    Name = soundName
                }) { Id = string.Empty });
            return new Shape(
                new NonVisualShapeProperties(properties,
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 100000, Y = 100000 },
                        new A.Extents { Cx = 500000, Cy = 300000 }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    }),
                new TextBody(new A.BodyProperties(), new A.ListStyle(),
                    new A.Paragraph(new A.EndParagraphRunProperties())));
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
