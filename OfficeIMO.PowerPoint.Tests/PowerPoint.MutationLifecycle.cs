using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public sealed class PowerPointMutationLifecycleTests {
        [Fact]
        public void RemovingOnlyCustomShowSlideRemovesShowAndActions() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointAutoShape actionShape = source.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            action.Action = "PPAction://CustomShow?id=17&return=true";

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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void ImportingMutuallyLinkedSlidesImportsListedTargetsOnce(bool batch, bool sameDocument) {
            using PowerPointDocument source =
                PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointTextRun firstRun = first.AddTextBox("Second")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide second = source.AddSlide();
            PowerPointTextRun secondRun = second.AddTextBox("First")
                .Paragraphs.Single().Runs.Single();
            firstRun.SetHyperlink(second);
            secondRun.SetHyperlink(first);
            using PowerPointDocument target =
                PowerPointDocument.Create();

            PowerPointDocument destination = sameDocument ? source : target;
            int offset = destination.Slides.Count;
            PowerPointSlide importedFirst;
            if (batch) {
                var imported = destination.ImportSlides(source, new[] { 0, 1, 0 });
                Assert.Equal(2, imported.Count);
                importedFirst = imported[0];
                Assert.Same(destination.Slides[offset + 1], imported[1]);
            } else importedFirst = destination.ImportSlide(source, 0);

            Assert.Equal(offset + 2, destination.Slides.Count);
            PowerPointSlide importedSecond = destination.Slides[offset + 1];
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
            Assert.Empty(destination.ValidateDocument());
            using var reopened = PowerPointDocument.Load(new MemoryStream(destination.ToBytes()));
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(offset + 2, reopened.Slides.Count);
        }

        [Fact]
        public void ImportAndExportPreserveNotesAlongsideAudioRelationships() {
            using PowerPointDocument source =
                PowerPointDocument.Create();
            PowerPointSlide sourceSlide = source.AddSlide();
            using (var audio = new MemoryStream(CreateWave(),
                       writable: false)) {
                sourceSlide.AddAudio(audio, "audio/wav", ".wav");
            }
            sourceSlide.Notes.Text = "Media notes";
            using PowerPointDocument target =
                PowerPointDocument.Create();

            PowerPointSlide imported = target.ImportSlide(source, 0);

            Assert.Equal("Media notes", imported.Notes.Text);
            Assert.Single(imported.Media);
            Assert.Empty(target.ValidateDocument());

            using var exportedStream = new MemoryStream();
            source.ExportSlide(0, exportedStream);
            exportedStream.Position = 0;
            using PowerPointDocument exported =
                PowerPointDocument.Load(exportedStream);
            PowerPointSlide exportedSlide = Assert.Single(exported.Slides);
            Assert.Equal("Media notes", exportedSlide.Notes.Text);
            Assert.Single(exportedSlide.Media);
            Assert.Empty(exported.ValidateDocument());
        }

        [Fact]
        public void RemovingSlideCleansLayoutAndSoundedInboundLinks() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide source = presentation.AddSlide(
                PowerPointSlideLayoutType.Title);
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            SlideLayoutPart layout = slide.SlidePart.SlideLayoutPart!;
            string layoutRelationshipId = slide.SlidePart.GetIdOfPart(layout);
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            using var input = new MemoryStream(original, writable: false);
            PowerPointMedia first = slide.AddAudio(input, "audio/wav",
                ".wav");
            first.Name = first.MediaReferenceId!;
            string originalName = first.Name!;
            PowerPointMedia second = Assert.IsType<PowerPointMedia>(
                first.Duplicate(offsetX: OfficeIMO.OfficeLength.Emus(1000000)));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            using PowerPointDocument source =
                PowerPointDocument.Create();
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
            using PowerPointDocument target =
                PowerPointDocument.Create();

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
            using PowerPointDocument source =
                PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointTextRun link = first.AddTextBox("Broken target")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide brokenTarget = source.AddSlide();
            link.SetHyperlink(brokenTarget);
            brokenTarget.SlidePart.DeletePart(
                brokenTarget.SlidePart.SlideLayoutPart!);
            using PowerPointDocument target =
                PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide first = presentation.AddSlide();
            PowerPointAutoShape local = first.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
            first.AddClassicAnimation(local,
                PowerPointClassicAnimationEffect.Fade);
            PowerPointSlide second = presentation.AddSlide();
            PowerPointAutoShape foreign = second.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
            Assert.Equal(local.Id, foreign.Id);

            Assert.Throws<ArgumentException>(() =>
                first.RemoveClassicAnimation(foreign));

            Assert.Single(first.ClassicAnimations);
            Assert.Empty(second.ClassicAnimations);
        }

        [Fact]
        public void ClearingClassicAnimationsPreservesMediaTiming() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadsClassicAnimationsIgnoringAdvancedEffectsInsideExtensions(bool keepAuthoredMetadata) {
            using var presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox target = slide.AddTextBox("Classic animation");
            slide.AddClassicAnimation(target, PowerPointClassicAnimationEffect.Fade, new PowerPointClassicAnimationOptions {
                Automatic = true, DelayMilliseconds = 400,
                BuildType = PowerPointClassicAnimationBuildType.ByLevel2Paragraph, Reverse = true, AnimateBackground = true
            });
            using var bytes = new MemoryStream(presentation.ToBytes());
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(bytes, leaveOpen: true);
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace producer = "https://schemas.officeimo.net/powerpoint/2026/classic-animations";
            const string part = "/ppt/slides/slide1.xml";
            XDocument markup = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(part)));
            if (!keepAuthoredMetadata) markup.Root!.Element(p + "extLst")!.Elements(p + "ext")
                .Single(extension => extension.Element(producer + "classicAnimations") != null).Remove();
            markup.Root!.Element(p + "timing")!.Add(new XElement(p + "extLst",
                new XElement(p + "ext", new XAttribute("uri", "{436DBF58-C5C2-449F-93DA-8926A8FAB706}"),
                    new XElement(p + "animMotion", new XElement(p + "cBhvr",
                        new XElement(p + "tgtEl", new XElement(p + "spTgt", new XAttribute("spid", target.Id!.Value))))))));
            byte[] modified = Encoding.UTF8.GetBytes(markup.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(part, modified, modified.Length);
            using var input = new MemoryStream();
            package.Save(input);
            input.Position = 0;

            using var reloaded = PowerPointDocument.Load(input);
            PowerPointClassicAnimation animation = Assert.Single(reloaded.Slides.Single().ClassicAnimations);
            Assert.Equal(target.Id, animation.ShapeId);
            Assert.Equal(PowerPointClassicAnimationEffect.Fade, animation.Effect);
            Assert.Equal(PowerPointClassicAnimationBuildType.ByLevel2Paragraph, animation.BuildType);
            Assert.Equal(400, animation.DelayMilliseconds);
            Assert.True(animation.Automatic && animation.Reverse && animation.AnimateBackground);
            Assert.False(reloaded.IsSdkLoaded);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void NativeClassicBuildInferencePreservesSdkHandlingOfInvalidTokensAndComments(bool addComment) {
            using var source = PowerPointDocument.Create();
            PowerPointSlide slide = source.AddSlide();
            slide.AddClassicAnimation(slide.AddTextBox("Classic animation"), PowerPointClassicAnimationEffect.Fade);
            Slide root = slide.SlidePart.Slide!;
            root.SlideExtensionList!.Remove();
            BuildParagraph build = Assert.Single(root.Timing!.Descendants<BuildParagraph>());
            if (addComment) build.Parent!.Append(new OpenXmlMiscNode(XmlNodeType.Comment, "<!--producer notes-->"));
            else build.SetAttribute(new OpenXmlAttribute(string.Empty, "build", string.Empty, "unknown-build"));

            using var input = new MemoryStream(source.ToBytes());
            using var restored = PowerPointDocument.Load(input);
            PowerPointSlide restoredSlide = Assert.Single(restored.Slides);
            if (addComment) {
                Assert.Empty(restoredSlide.ClassicAnimations);
                Assert.True(restored.IsSdkLoaded);
            } else {
                Assert.Throws<FormatException>(() => _ = restoredSlide.ClassicAnimations);
                Assert.False(restored.IsSdkLoaded);
            }
        }

        [Fact]
        public void ClearingClassicAnimationsRemovesOnlyClassicMixedTiming() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
        public void ClearingClassicAnimationsKeepsOtherTypedActionsAttachedThroughSdkInterop() {
            using var presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddClassicAnimation(slide.AddTextBox("Classic animation"), PowerPointClassicAnimationEffect.Fade);
            Timing timing = slide.SlidePart.Slide!.Timing!;
            CommonTimeNode root = timing.Descendants<CommonTimeNode>().First();
            AnimateColor otherAction = root.AppendChild(new AnimateColor());

            slide.ClearClassicAnimations();

            Assert.Same(otherAction, Assert.Single(slide.SlidePart.Slide.Timing!.Descendants<AnimateColor>()));
            Assert.Empty(slide.ClassicAnimations);
        }

        [Fact]
        public void AuthoringAndClearingClassicAnimationsAndSoundKeepsThePackageOwned() {
            using var presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox shape = slide.AddTextBox("Classic animation");
            slide.AddClassicAnimation(shape, PowerPointClassicAnimationEffect.Fade, new PowerPointClassicAnimationOptions {
                Automatic = true, DelayMilliseconds = 400, AfterEffect = PowerPointClassicAnimationAfterEffect.Dim,
                RawDimColor = 0xFF336699U
            });
            byte[] wave = CreateWave();
            using var audio = new MemoryStream(wave, writable: false);
            slide.SetClassicAnimationSound(shape, audio, "Narration", stopExistingSounds: true);
            using var bytes = new MemoryStream(presentation.ToBytes());
            Assert.False(presentation.IsSdkLoaded);

            using var restored = PowerPointDocument.Load(bytes);
            PowerPointSlide restoredSlide = Assert.Single(restored.Slides);
            PowerPointTextBox restoredShape = Assert.Single(restoredSlide.TextBoxes);
            PowerPointClassicAnimation animation = Assert.Single(restoredSlide.ClassicAnimations);
            Assert.True(animation.PlaysSound && animation.StopsSound);
            Assert.Equal(PowerPointClassicAnimationAfterEffect.Dim, animation.AfterEffect);
            Assert.Equal(wave, restoredSlide.GetClassicAnimationSoundBytes(restoredShape));
            restoredSlide.ClearClassicAnimations();
            Assert.Empty(restoredSlide.ClassicAnimations);
            using var cleared = new MemoryStream(restored.ToBytes());
            Assert.False(restored.IsSdkLoaded);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(cleared);
            Assert.DoesNotContain(package.PartNames, name => name.StartsWith("/media/", StringComparison.Ordinal));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReplacingClassicTimingChecksForIdentifierExhaustionBeforeRemovingTheOldEffect(bool exhaustWithUnrelatedMedia) {
            using var document = PowerPointDocument.Create();
            PowerPointSlide slide = document.AddSlide();
            PowerPointTextBox shape = slide.AddTextBox("Classic animation");
            slide.AddClassicAnimation(shape, PowerPointClassicAnimationEffect.Fade);
            using var audio = new MemoryStream(CreateWave(), writable: false);
            slide.AddAudio(audio, "audio/wav", ".wav");
            Timing timing = slide.SlidePart.Slide!.Timing!;
            Audio retained = Assert.Single(timing.Descendants<Audio>());
            if (exhaustWithUnrelatedMedia) retained.CommonMediaNode!.CommonTimeNode!.Id = uint.MaxValue - 1;
            else Assert.Single(timing.Descendants<AnimateEffect>()).CommonBehavior!.CommonTimeNode!.Id = uint.MaxValue;
            string before = slide.SlidePart.Slide.OuterXml;

            if (exhaustWithUnrelatedMedia) {
                Assert.Throws<InvalidOperationException>(() => slide.SetClassicAnimationStopsSound(shape, true));
                Assert.Equal(before, slide.SlidePart.Slide.OuterXml);
            } else {
                slide.SetClassicAnimationStopsSound(shape, true);
                Assert.True(Assert.Single(slide.ClassicAnimations).StopsSound);
                Assert.Same(retained, Assert.Single(timing.Descendants<Audio>()));
            }
        }

        [Fact]
        public void RemovingOneClassicAnimationPreservesMixedMediaTiming() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
            PowerPointAutoShape second = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(700000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
        public void ReplacingRunHyperlinkPreservesSoundAndCleansOldTarget() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointSlide target = presentation.AddSlide();
            PowerPointTextRun run = source.AddTextBox("Target")
                .Paragraphs.Single().Runs.Single();
            byte[] soundBytes = CreateWave();
            using (var sound = new MemoryStream(soundBytes,
                       writable: false)) {
                run.SetClickSound(sound, "Old link sound");
            }
            run.SetClickStopsSound(true);
            run.SetHyperlink("https://example.test/old");
            Assert.Single(source.SlidePart.HyperlinkRelationships);
            Assert.Single(presentation.OpenXmlDocument.DataParts);
            Assert.Equal("Old link sound", run.ClickSoundName);
            Assert.Equal(soundBytes, run.GetClickSoundBytes());

            run.SetHyperlink(target);

            Assert.Empty(source.SlidePart.HyperlinkRelationships);
            Assert.Single(source.SlidePart.DataPartReferenceRelationships);
            Assert.Single(presentation.OpenXmlDocument.DataParts);
            Assert.Equal("Old link sound", run.ClickSoundName);
            Assert.Equal(soundBytes, run.GetClickSoundBytes());
            Assert.True(run.Run.RunProperties!
                .GetFirstChild<A.HyperlinkOnClick>()!.EndSound!.Value);
            Assert.Contains(source.SlidePart.Parts, pair => ReferenceEquals(
                pair.OpenXmlPart, target.SlidePart));

            run.ClearHyperlink();

            Assert.DoesNotContain(source.SlidePart.Parts, pair =>
                ReferenceEquals(pair.OpenXmlPart, target.SlidePart));
            Assert.Empty(source.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Equal(2, presentation.Slides.Count);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClearingRunHyperlinkIgnoresShapeNameMatchingRelationshipId() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextRun run = slide.AddTextBox("Link")
                .Paragraphs.Single().Runs.Single();
            run.SetHyperlink("https://example.test/remove");
            string relationshipId = Assert.Single(slide.SlidePart
                .HyperlinkRelationships).Id;
            slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000))
                .Name = relationshipId;

            run.ClearHyperlink();

            Assert.Empty(slide.SlidePart.HyperlinkRelationships);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void ClearingShapeSoundIgnoresShapeNameMatchingRelationshipId() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape sounded = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
            using (var sound = new MemoryStream(CreateWave(),
                       writable: false)) {
                sounded.SetClickSound(sound, "Remove sound");
            }
            string relationshipId = Assert.Single(slide.SlidePart
                .DataPartReferenceRelationships).Id;
            slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(700000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000))
                .Name = relationshipId;

            sounded.ClearClickSound();

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void FailedSoundIngestionDoesNotLeaveMediaParts() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(500000));
            using var sound = new MemoryStream(CreateWave(), writable: false);

            Assert.Throws<InvalidOperationException>(() =>
                slide.SetClassicAnimationSound(shape, sound, "No animation"));

            Assert.Empty(slide.SlidePart.DataPartReferenceRelationships);
            Assert.Empty(presentation.OpenXmlDocument.DataParts);
        }

        [Fact]
        public void FailedMediaReplacementPreservesExistingPayload() {
            byte[] original = CreateWave();
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
