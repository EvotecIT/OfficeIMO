using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptTransitionTests {
        [Theory]
        [InlineData(SlideTransition.Cut, 0, 1)]
        [InlineData(SlideTransition.Fade, 6, 0)]
        [InlineData(SlideTransition.Wipe, 10, 0)]
        [InlineData(SlideTransition.BlindsVertical, 2, 0)]
        [InlineData(SlideTransition.BlindsHorizontal, 2, 1)]
        [InlineData(SlideTransition.CombHorizontal, 21, 0)]
        [InlineData(SlideTransition.CombVertical, 21, 1)]
        [InlineData(SlideTransition.PushLeft, 20, 0)]
        [InlineData(SlideTransition.PushUp, 20, 1)]
        [InlineData(SlideTransition.PushRight, 20, 2)]
        [InlineData(SlideTransition.PushDown, 20, 3)]
        public void NativeWriter_AuthorsSupportedTransitionAndAdvanceSettings(
            SlideTransition transition, byte effectType, byte effectDirection) {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                slide.Transition = transition;
                slide.TransitionSpeed = SlideTransitionSpeed.Fast;
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
            Assert.Equal(SlideTransitionSpeed.Fast, projectedSlide.TransitionSpeed);
            Assert.False(projectedSlide.TransitionAdvanceOnClick);
            Assert.Equal(4.25, projectedSlide.TransitionAdvanceAfterSeconds);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksModernTransitionWithoutLegacyEquivalent() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide();
            slide.Transition = SlideTransition.Morph;

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
                slide.Transition = SlideTransition.Fade;
                slide.TransitionSpeed = SlideTransitionSpeed.Slow;
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                slide.Transition = SlideTransition.PushDown;
                slide.TransitionSpeed = SlideTransitionSpeed.Fast;
                slide.TransitionAdvanceOnClick = false;
                slide.TransitionAdvanceAfterSeconds = 7.5;

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptTransition transition = Assert.IsType<LegacyPptTransition>(
                Assert.Single(saved.Slides).Transition);
            Assert.Equal(LegacyPptTransitionEffect.Push, transition.Effect);
            Assert.Equal(3, transition.EffectDirection);
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
