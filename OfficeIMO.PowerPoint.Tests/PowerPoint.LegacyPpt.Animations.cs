using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptAnimationTests {
        [Fact]
        public void NativeWriter_AuthorsProjectsAndInspectsClassicAnimations() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape rectangle = slide.AddRectangle(
                    100000, 100000, 1200000, 500000);
                PowerPointTextBox text = slide.AddTextBox("Animated text",
                    100000, 800000, 1800000, 500000);
                slide.AddClassicAnimation(rectangle,
                    PowerPointClassicAnimationEffect.Wipe,
                    new PowerPointClassicAnimationOptions {
                        Direction = 2,
                        Reverse = true
                    });
                slide.AddClassicAnimation(text,
                    PowerPointClassicAnimationEffect.Fly,
                    new PowerPointClassicAnimationOptions {
                        Direction = 0,
                        BuildType = PowerPointClassicAnimationBuildType.ByLevel2Paragraph,
                        Automatic = true,
                        DelayMilliseconds = 250,
                        AnimateBackground = true,
                        AfterEffect = PowerPointClassicAnimationAfterEffect.Dim,
                        TextBuild = PowerPointClassicTextBuild.ByWord,
                        RawDimColor = 0x08000002U
                    });

                Assert.Equal(2, slide.ClassicAnimations.Count);
                LegacyPptWritePreflightReport preflight = source.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                var sourceErrors = source.ValidateDocument();
                Assert.True(sourceErrors.Count == 0, string.Join(Environment.NewLine,
                    sourceErrors.Select(error => error.Description + " | "
                        + error.Path?.XPath)) + Environment.NewLine
                    + slide.SlidePart.Slide!.OuterXml);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, legacy.CreateImportReport().AnimationCount);
            LegacyPptAnimation first = Assert.IsType<LegacyPptAnimation>(
                legacy.Slides[0].Shapes[0].Animation);
            Assert.Equal(LegacyPptAnimationEffect.Wipe, first.Effect);
            Assert.Equal(2, first.EffectDirection);
            Assert.True(first.Reverse);
            Assert.False(first.Automatic);
            Assert.Equal(0, first.RawUnused);
            Assert.False(first.HasSoundOverride);
            LegacyPptAnimation second = Assert.IsType<LegacyPptAnimation>(
                legacy.Slides[0].Shapes[1].Animation);
            Assert.Equal(LegacyPptAnimationEffect.Fly, second.Effect);
            Assert.Equal(LegacyPptAnimationBuildType.ByLevel2Paragraph,
                second.BuildType);
            Assert.True(second.Automatic);
            Assert.Equal(250, second.DelayMilliseconds);
            Assert.True(second.AnimateBackground);
            Assert.Equal(LegacyPptAnimationAfterEffect.Dim,
                second.AfterEffect);
            Assert.Equal(LegacyPptTextBuildSubEffect.ByWord,
                second.TextBuildSubEffect);
            Assert.Equal(0x08000002U, second.RawDimColor);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            Assert.Equal(2, projected.Slides[0].ClassicAnimations.Count);
            PowerPointClassicAnimation projectedSecond = projected.Slides[0]
                .ClassicAnimations[1];
            Assert.Equal(PowerPointClassicAnimationEffect.Fly,
                projectedSecond.Effect);
            Assert.Equal(PowerPointClassicAnimationBuildType.ByLevel2Paragraph,
                projectedSecond.BuildType);
            Assert.True(projectedSecond.Automatic);
            Assert.Equal(250, projectedSecond.DelayMilliseconds);
            Assert.Equal(PowerPointClassicAnimationAfterEffect.Dim,
                projectedSecond.AfterEffect);
            Assert.Equal(PowerPointClassicTextBuild.ByWord,
                projectedSecond.TextBuild);
            Assert.True(projected.InspectAnimations().HasAnimations);
            var projectedErrors = projected.ValidateDocument();
            Assert.True(projectedErrors.Count == 0, string.Join(Environment.NewLine,
                projectedErrors.Select(error => error.Description + " | "
                    + error.Path?.XPath)));
        }

        [Fact]
        public void PublicApi_RejectsEffectDirectionOutsideLegacyContract() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1200000, 500000);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddClassicAnimation(shape,
                    PowerPointClassicAnimationEffect.Wheel,
                    new PowerPointClassicAnimationOptions { Direction = 7 }));
            Assert.Empty(slide.ClassicAnimations);
        }

        [Fact]
        public void ImportedAnimation_UnrelatedShapeEditPreservesExactSemantics() {
            byte[] sourceBytes = CreateBinaryAnimationPresentation();
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
            LegacyPptAnimation animation = Assert.IsType<LegacyPptAnimation>(
                saved.Slides[0].Shapes[0].Animation);
            Assert.Equal(LegacyPptAnimationEffect.Wipe, animation.Effect);
            Assert.Equal(2, animation.EffectDirection);
            Assert.True(animation.Reverse);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedAnimation_EditRemoveAndAddUseIncrementalRecords() {
            byte[] sourceBytes = CreateBinaryAnimationPresentation();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                PowerPointShape shape = slide.Shapes[0];
                Assert.True(slide.RemoveClassicAnimation(shape));
                slide.AddClassicAnimation(shape,
                    PowerPointClassicAnimationEffect.Wheel,
                    new PowerPointClassicAnimationOptions {
                        Direction = 4,
                        Automatic = true,
                        DelayMilliseconds = 700,
                        AfterEffect = PowerPointClassicAnimationAfterEffect.HideImmediately
                    });
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation edited = LegacyPptPresentation.Load(editedBytes);
            LegacyPptAnimation animation = Assert.IsType<LegacyPptAnimation>(
                edited.Slides[0].Shapes[0].Animation);
            Assert.Equal(LegacyPptAnimationEffect.Wheel, animation.Effect);
            Assert.Equal(4, animation.EffectDirection);
            Assert.True(animation.Automatic);
            Assert.Equal(700, animation.DelayMilliseconds);
            Assert.Equal(LegacyPptAnimationAfterEffect.HideImmediately,
                animation.AfterEffect);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                Assert.True(imported.Slides[0].RemoveClassicAnimation(
                    imported.Slides[0].Shapes[0]));
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            Assert.Null(LegacyPptPresentation.Load(removedBytes)
                .Slides[0].Shapes[0].Animation);
        }

        [Fact]
        public void ImportedSlide_CanAddFirstClassicAnimationIncrementally() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                    100000, 100000, 1200000, 500000);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                slide.AddClassicAnimation(slide.Shapes[0],
                    PowerPointClassicAnimationEffect.Fade);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            Assert.Equal(LegacyPptAnimationEffect.Fade,
                Assert.IsType<LegacyPptAnimation>(LegacyPptPresentation
                    .Load(savedBytes).Slides[0].Shapes[0].Animation).Effect);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void AnimationSound_AuthorsProjectsAndReplacesIncrementally() {
            byte[] wave = CreateWavePayload();
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1200000, 500000);
                slide.AddClassicAnimation(shape,
                    PowerPointClassicAnimationEffect.Fade);
                using var audio = new MemoryStream(wave, writable: false);
                slide.SetClassicAnimationSound(shape, audio,
                    "Animation Chime", stopExistingSounds: true);
                Assert.Equal(wave, slide.GetClassicAnimationSoundBytes(shape));
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptSound originalSound = Assert.Single(original.Sounds);
            LegacyPptAnimation originalAnimation = Assert.IsType<LegacyPptAnimation>(
                original.Slides[0].Shapes[0].Animation);
            Assert.True(originalAnimation.PlaysSound);
            Assert.True(originalAnimation.StopsSound);
            Assert.Equal(originalSound.Id, originalAnimation.SoundIdReference);
            Assert.Equal(1, original.CreateImportReport().AnimationSoundCount);

            byte[] replacement = wave.Concat(new byte[] { 0xA1, 0xA2 }).ToArray();
            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                PowerPointShape shape = slide.Shapes[0];
                PowerPointClassicAnimation projected =
                    Assert.Single(slide.ClassicAnimations);
                Assert.Equal("Animation Chime", projected.SoundName);
                Assert.True(projected.PlaysSound);
                Assert.True(projected.StopsSound);
                Assert.Equal(wave, slide.GetClassicAnimationSoundBytes(shape));
                using var audio = new MemoryStream(replacement, writable: false);
                slide.SetClassicAnimationSound(shape, audio,
                    "Replacement Animation", stopExistingSounds: false);
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal(2, saved.Sounds.Count);
            LegacyPptSound appended = saved.Sounds.Single(sound =>
                sound.Name == "Replacement Animation");
            Assert.Equal(replacement, appended.DataBytes);
            LegacyPptAnimation savedAnimation = Assert.IsType<LegacyPptAnimation>(
                saved.Slides[0].Shapes[0].Animation);
            Assert.True(savedAnimation.PlaysSound);
            Assert.False(savedAnimation.StopsSound);
            Assert.Equal(appended.Id, savedAnimation.SoundIdReference);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void PptxTiming_MaterializesClassicSoundsAndAfterEffects() {
            byte[] wave = CreateWavePayload();
            byte[] pptx;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape dimmed = slide.AddRectangle(
                    100000, 100000, 800000, 400000);
                PowerPointAutoShape hiddenImmediately = slide.AddRectangle(
                    100000, 600000, 800000, 400000);
                PowerPointAutoShape hiddenOnClick = slide.AddRectangle(
                    100000, 1100000, 800000, 400000);
                slide.AddClassicAnimation(dimmed,
                    PowerPointClassicAnimationEffect.Fade,
                    new PowerPointClassicAnimationOptions {
                        AfterEffect = PowerPointClassicAnimationAfterEffect.Dim,
                        RawDimColor = 0xFE332211U,
                        StopsSound = true
                    });
                slide.AddClassicAnimation(hiddenImmediately,
                    PowerPointClassicAnimationEffect.Fade,
                    new PowerPointClassicAnimationOptions {
                        AfterEffect = PowerPointClassicAnimationAfterEffect
                            .HideImmediately
                    });
                slide.AddClassicAnimation(hiddenOnClick,
                    PowerPointClassicAnimationEffect.Fade,
                    new PowerPointClassicAnimationOptions {
                        AfterEffect = PowerPointClassicAnimationAfterEffect
                            .HideOnNextClick
                    });
                using var audio = new MemoryStream(wave, writable: false);
                slide.SetClassicAnimationSound(dimmed, audio,
                    "Animation Chime", stopExistingSounds: true);

                P.Timing timing = Assert.IsType<P.Timing>(
                    slide.SlidePart.Slide!.Timing);
                P.Audio sound = Assert.Single(timing.Descendants<P.Audio>());
                P.SoundTarget soundTarget = Assert.IsType<P.SoundTarget>(
                    sound.CommonMediaNode!.GetFirstChild<P.TargetElement>()!
                        .FirstChild);
                Assert.Equal(slide.ClassicAnimations[0].SoundRelationshipId,
                    soundTarget.Embed?.Value);
                P.Command stopSound = Assert.Single(
                    timing.Descendants<P.Command>());
                Assert.Equal(P.CommandValues.Event,
                    stopSound.Type?.Value);
                Assert.Equal("onstopaudio",
                    stopSound.CommandName?.Value);
                Assert.NotNull(stopSound.CommonBehavior?
                    .GetFirstChild<P.TargetElement>()?
                    .GetFirstChild<P.SlideTarget>());
                Assert.Contains(stopSound.Descendants<P.Condition>(),
                    condition => condition.Event?.Value ==
                        P.TriggerEventValues.Begin);
                P.AnimateColor dim = Assert.Single(
                    timing.Descendants<P.AnimateColor>());
                Assert.Equal(dimmed.Id!.Value.ToString(), dim.CommonBehavior!
                    .GetFirstChild<P.TargetElement>()!
                    .GetFirstChild<P.ShapeTarget>()!.ShapeId!.Value);
                Assert.Equal("112233", dim.ToColor!
                    .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
                P.SetBehavior[] hidden = timing.Descendants<P.SetBehavior>()
                    .Where(set => string.Equals(set.ToVariantValue?
                            .GetFirstChild<P.StringVariantValue>()?.Val?.Value,
                        "hidden", StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                Assert.Equal(2, hidden.Length);
                Assert.Contains(hidden, set => set.Descendants<P.Condition>()
                    .Any(condition => condition.Event?.Value ==
                        P.TriggerEventValues.End));
                Assert.Contains(hidden, set => set.Descendants<P.Condition>()
                    .Any(condition => condition.Event?.Value ==
                        P.TriggerEventValues.OnClick
                        && condition.Descendants<P.SlideTarget>().Any()));
                Assert.True(slide.HasOnlyClassicAnimationTiming());
                Assert.Empty(source.ValidateDocument());
                pptx = source.ToBytes();
            }

            using var input = new MemoryStream(pptx, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            Assert.Equal(3, reopened.Slides[0].ClassicAnimations.Count);
            Assert.True(reopened.Slides[0].ClassicAnimations[0].StopsSound);
            Assert.Equal(wave, reopened.Slides[0]
                .GetClassicAnimationSoundBytes(
                    reopened.Slides[0].Shapes[0]));
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void PptxTiming_MultipleStopSoundCommandsRemainClassicForPptExport() {
            byte[] ppt;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape first = slide.AddRectangle(
                    100000, 100000, 800000, 400000);
                PowerPointAutoShape second = slide.AddRectangle(
                    100000, 600000, 800000, 400000);
                var options = new PowerPointClassicAnimationOptions {
                    StopsSound = true
                };
                slide.AddClassicAnimation(first,
                    PowerPointClassicAnimationEffect.Fade, options);
                slide.AddClassicAnimation(second,
                    PowerPointClassicAnimationEffect.Cut, options);

                Assert.Equal(2, slide.SlidePart.Slide!.Timing!
                    .Descendants<P.Command>().Count());
                Assert.True(slide.HasOnlyClassicAnimationTiming());
                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                Assert.Empty(source.ValidateDocument());
                ppt = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptAnimation[] animations = LegacyPptPresentation.Load(ppt)
                .Slides[0].Shapes.Select(shape => shape.Animation)
                .OfType<LegacyPptAnimation>().ToArray();
            Assert.Equal(2, animations.Length);
            Assert.All(animations, animation =>
                Assert.True(animation.StopsSound));
        }

        [Fact]
        public void NativeWriter_ConvertsPowerPointAuthoredClassicTiming() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointTextBox shape = slide.AddTextBox("One\nTwo\nThree",
                    100000, 100000, 1800000, 900000);
                uint shapeId = Assert.IsType<uint>(shape.Id);
                slide.SlidePart.Slide!.Append(
                    CreatePowerPointAuthoredClassicTiming(shapeId));

                PowerPointClassicAnimation animation = Assert.Single(
                    slide.ClassicAnimations);
                Assert.Equal(PowerPointClassicAnimationEffect.Wipe,
                    animation.Effect);
                Assert.Equal(3, animation.Direction);
                Assert.Equal(PowerPointClassicAnimationBuildType.ByLevel2Paragraph,
                    animation.BuildType);
                Assert.False(animation.Automatic);
                Assert.Empty(slide.SlidePart.Slide
                    .GetFirstChild<P.SlideExtensionList>()?
                    .Elements<P.SlideExtension>()
                    ?? Enumerable.Empty<P.SlideExtension>());

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                PowerPointFeatureReport features = source.InspectFeatures();
                PowerPointFeatureFinding classic = Assert.Single(
                    features.FindFeatures("Classic animations"));
                Assert.Equal(PowerPointFeatureSupportLevel.Editable,
                    classic.SupportLevel);
                Assert.Equal(1, classic.Count);
                Assert.Empty(features.FindFeatures("Animations and timing"));
                Assert.Empty(source.ValidateDocument());
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptAnimation animationAtom = Assert.IsType<LegacyPptAnimation>(
                LegacyPptPresentation.Load(bytes).Slides[0].Shapes[0].Animation);
            Assert.Equal(LegacyPptAnimationEffect.Wipe,
                animationAtom.Effect);
            Assert.Equal(3, animationAtom.EffectDirection);
            Assert.Equal(LegacyPptAnimationBuildType.ByLevel2Paragraph,
                animationAtom.BuildType);
        }

        [Fact]
        public void NativeWriter_BlocksAdvancedTimingMixedWithClassicAnimation() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1200000, 500000);
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Fade);
            slide.SlidePart.Slide!.Timing!.Descendants<P.ChildTimeNodeList>()
                .Last().Append(new P.AnimateMotion());

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-TIMING"
                && finding.Feature == LegacyPptFeature.Animations);
        }

        [Fact]
        public void NativeWriter_ConvertsAfterPreviousButBlocksWithPrevious() {
            byte[] automaticBytes;
            using (PowerPointPresentation automatic =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = automatic.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1200000, 500000);
                P.Timing timing = CreatePowerPointAuthoredClassicTiming(
                    Assert.IsType<uint>(shape.Id));
                P.CommonTimeNode effectOwner = timing
                    .Descendants<P.CommonTimeNode>().Single(node =>
                        node.Id?.Value == 5U);
                effectOwner.NodeType = P.TimeNodeValues.AfterEffect;
                effectOwner.GetFirstChild<P.StartConditionList>()!
                    .GetFirstChild<P.Condition>()!.Delay = "325";
                slide.SlidePart.Slide!.Append(timing);

                PowerPointClassicAnimation inferred = Assert.Single(
                    slide.ClassicAnimations);
                Assert.True(inferred.Automatic);
                Assert.Equal(325, inferred.DelayMilliseconds);
                Assert.True(automatic.AnalyzeLegacyPptWrite().CanWrite);
                automaticBytes = automatic.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptAnimation atom = Assert.IsType<LegacyPptAnimation>(
                LegacyPptPresentation.Load(automaticBytes).Slides[0]
                    .Shapes[0].Animation);
            Assert.True(atom.Automatic);
            Assert.Equal(325, atom.DelayMilliseconds);

            using PowerPointPresentation concurrent =
                PowerPointPresentation.Create();
            PowerPointSlide concurrentSlide = concurrent.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointAutoShape concurrentShape = concurrentSlide.AddRectangle(
                100000, 100000, 1200000, 500000);
            P.Timing concurrentTiming = CreatePowerPointAuthoredClassicTiming(
                Assert.IsType<uint>(concurrentShape.Id));
            concurrentTiming.Descendants<P.CommonTimeNode>().Single(node =>
                    node.Id?.Value == 5U).NodeType =
                P.TimeNodeValues.WithEffect;
            concurrentSlide.SlidePart.Slide!.Append(concurrentTiming);

            Assert.Empty(concurrentSlide.ClassicAnimations);
            Assert.Contains(concurrent.AnalyzeLegacyPptWrite().Findings,
                finding => finding.Code == "PPT-WRITE-TIMING");
        }

        private static byte[] CreateBinaryAnimationPresentation() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1200000, 500000);
            slide.AddClassicAnimation(shape,
                PowerPointClassicAnimationEffect.Wipe,
                new PowerPointClassicAnimationOptions {
                    Direction = 2,
                    Reverse = true
                });
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

        private static P.Timing CreatePowerPointAuthoredClassicTiming(
            uint shapeId) => new($$"""
            <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:tnLst>
                <p:par>
                  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                    <p:childTnLst>
                      <p:seq concurrent="1" nextAc="seek">
                        <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                          <p:childTnLst>
                            <p:par><p:cTn id="3" fill="hold">
                              <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                              <p:childTnLst><p:par><p:cTn id="4" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst><p:par><p:cTn id="5" presetID="22" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="clickEffect">
                                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                  <p:endCondLst><p:cond evt="begin" delay="0"><p:tn val="5"/></p:cond></p:endCondLst>
                                  <p:childTnLst>
                                    <p:set>
                                      <p:cBhvr>
                                        <p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                        <p:tgtEl><p:spTgt spid="{{shapeId}}"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl>
                                        <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                                      </p:cBhvr>
                                      <p:to><p:strVal val="visible"/></p:to>
                                    </p:set>
                                    <p:animEffect transition="in" filter="wipe(down)">
                                      <p:cBhvr>
                                        <p:cTn id="7" dur="500"/>
                                        <p:tgtEl><p:spTgt spid="{{shapeId}}"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl>
                                      </p:cBhvr>
                                    </p:animEffect>
                                  </p:childTnLst>
                                </p:cTn></p:par></p:childTnLst>
                              </p:cTn></p:par></p:childTnLst>
                            </p:cTn></p:par>
                          </p:childTnLst>
                        </p:cTn>
                        <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                        <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                      </p:seq>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:tnLst>
              <p:bldLst><p:bldP spid="{{shapeId}}" grpId="0" build="p" bldLvl="2"/></p:bldLst>
            </p:timing>
            """);
    }
}
