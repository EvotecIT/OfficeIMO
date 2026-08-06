using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptTransitionTests {
        [Theory]
        [InlineData(PowerPointSlideTransition.Cut, 0, 0)]
        [InlineData(PowerPointSlideTransition.CutThroughBlack, 0, 1)]
        [InlineData(PowerPointSlideTransition.Random, 1, 0)]
        [InlineData(PowerPointSlideTransition.Fade, 23, 0)]
        [InlineData(PowerPointSlideTransition.FadeThroughBlack, 6, 0)]
        [InlineData(PowerPointSlideTransition.Wipe, 10, 0)]
        [InlineData(PowerPointSlideTransition.WipeUp, 10, 1)]
        [InlineData(PowerPointSlideTransition.WipeRight, 10, 2)]
        [InlineData(PowerPointSlideTransition.WipeDown, 10, 3)]
        [InlineData(PowerPointSlideTransition.BlindsVertical, 2, 0)]
        [InlineData(PowerPointSlideTransition.BlindsHorizontal, 2, 1)]
        [InlineData(PowerPointSlideTransition.CheckerHorizontal, 3, 0)]
        [InlineData(PowerPointSlideTransition.CheckerVertical, 3, 1)]
        [InlineData(PowerPointSlideTransition.CoverLeft, 4, 0)]
        [InlineData(PowerPointSlideTransition.CoverUp, 4, 1)]
        [InlineData(PowerPointSlideTransition.CoverRight, 4, 2)]
        [InlineData(PowerPointSlideTransition.CoverDown, 4, 3)]
        [InlineData(PowerPointSlideTransition.CoverLeftUp, 4, 4)]
        [InlineData(PowerPointSlideTransition.CoverRightUp, 4, 5)]
        [InlineData(PowerPointSlideTransition.CoverLeftDown, 4, 6)]
        [InlineData(PowerPointSlideTransition.CoverRightDown, 4, 7)]
        [InlineData(PowerPointSlideTransition.Dissolve, 5, 0)]
        [InlineData(PowerPointSlideTransition.UncoverLeft, 7, 0)]
        [InlineData(PowerPointSlideTransition.UncoverUp, 7, 1)]
        [InlineData(PowerPointSlideTransition.UncoverRight, 7, 2)]
        [InlineData(PowerPointSlideTransition.UncoverDown, 7, 3)]
        [InlineData(PowerPointSlideTransition.UncoverLeftUp, 7, 4)]
        [InlineData(PowerPointSlideTransition.UncoverRightUp, 7, 5)]
        [InlineData(PowerPointSlideTransition.UncoverLeftDown, 7, 6)]
        [InlineData(PowerPointSlideTransition.UncoverRightDown, 7, 7)]
        [InlineData(PowerPointSlideTransition.RandomBarsHorizontal, 8, 0)]
        [InlineData(PowerPointSlideTransition.RandomBarsVertical, 8, 1)]
        [InlineData(PowerPointSlideTransition.StripsLeftUp, 9, 4)]
        [InlineData(PowerPointSlideTransition.StripsRightUp, 9, 5)]
        [InlineData(PowerPointSlideTransition.StripsLeftDown, 9, 6)]
        [InlineData(PowerPointSlideTransition.StripsRightDown, 9, 7)]
        [InlineData(PowerPointSlideTransition.BoxOut, 11, 0)]
        [InlineData(PowerPointSlideTransition.BoxIn, 11, 1)]
        [InlineData(PowerPointSlideTransition.SplitHorizontalOut, 13, 0)]
        [InlineData(PowerPointSlideTransition.SplitHorizontalIn, 13, 1)]
        [InlineData(PowerPointSlideTransition.SplitVerticalOut, 13, 2)]
        [InlineData(PowerPointSlideTransition.SplitVerticalIn, 13, 3)]
        [InlineData(PowerPointSlideTransition.Diamond, 17, 0)]
        [InlineData(PowerPointSlideTransition.Plus, 18, 0)]
        [InlineData(PowerPointSlideTransition.Wedge, 19, 0)]
        [InlineData(PowerPointSlideTransition.CombHorizontal, 21, 0)]
        [InlineData(PowerPointSlideTransition.CombVertical, 21, 1)]
        [InlineData(PowerPointSlideTransition.PushLeft, 20, 0)]
        [InlineData(PowerPointSlideTransition.PushUp, 20, 1)]
        [InlineData(PowerPointSlideTransition.PushRight, 20, 2)]
        [InlineData(PowerPointSlideTransition.PushDown, 20, 3)]
        [InlineData(PowerPointSlideTransition.Newsflash, 22, 0)]
        [InlineData(PowerPointSlideTransition.WheelOneSpoke, 26, 1)]
        [InlineData(PowerPointSlideTransition.WheelTwoSpokes, 26, 2)]
        [InlineData(PowerPointSlideTransition.WheelThreeSpokes, 26, 3)]
        [InlineData(PowerPointSlideTransition.WheelFourSpokes, 26, 4)]
        [InlineData(PowerPointSlideTransition.WheelEightSpokes, 26, 8)]
        [InlineData(PowerPointSlideTransition.Circle, 27, 0)]
        public void NativeWriter_AuthorsSupportedTransitionAndAdvanceSettings(
            PowerPointSlideTransition transition, byte effectType, byte effectDirection) {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Transition = transition;
                slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Fast;
                slide.TransitionAdvanceOnClick = false;
                slide.TransitionAdvanceAfterSeconds = 4.25;

                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptTransition binary = Assert.IsType<LegacyPptTransition>(
                Assert.Single(LegacyPptPresentation.Load(bytes).Slides).Transition);
            Assert.Equal(effectType, binary.RawEffectType);
            Assert.Equal(effectDirection, binary.EffectDirection);
            Assert.Equal(2, binary.Speed);
            Assert.False(binary.ManualAdvance);
            Assert.True(binary.AutoAdvance);
            Assert.Equal(4250, binary.SlideTimeMilliseconds);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            PowerPointSlide projectedSlide = projected.Slides[0];
            Assert.Equal(transition, projectedSlide.Transition);
            Assert.Equal(PowerPointSlideTransitionSpeed.Fast, projectedSlide.TransitionSpeed);
            Assert.False(projectedSlide.TransitionAdvanceOnClick);
            Assert.Equal(4.25, projectedSlide.TransitionAdvanceAfterSeconds);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksModernTransitionWithoutLegacyEquivalent() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide();
            slide.Transition = PowerPointSlideTransition.Morph;

            LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();

            LegacyPptWriteFinding finding = Assert.Single(report.Findings,
                item => item.Code == "PPT-WRITE-TRANSITION");
            Assert.Contains("no PowerPoint 97-2003 representation",
                finding.Description, StringComparison.Ordinal);
        }

        [Fact]
        public void ImportedTransitionEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Transition = PowerPointSlideTransition.Fade;
                slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Slow;
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                slide.Transition = PowerPointSlideTransition.WheelEightSpokes;
                slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Fast;
                slide.TransitionAdvanceOnClick = false;
                slide.TransitionAdvanceAfterSeconds = 7.5;

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptTransition transition = Assert.IsType<LegacyPptTransition>(
                Assert.Single(saved.Slides).Transition);
            Assert.Equal(LegacyPptTransitionEffect.Wheel, transition.Effect);
            Assert.Equal(8, transition.EffectDirection);
            Assert.Equal(2, transition.Speed);
            Assert.False(transition.ManualAdvance);
            Assert.True(transition.AutoAdvance);
            Assert.Equal(7500, transition.SlideTimeMilliseconds);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }
    }
}
