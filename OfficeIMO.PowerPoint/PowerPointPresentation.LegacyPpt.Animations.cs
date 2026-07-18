using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyAnimations(PowerPointSlide slide,
            IEnumerable<LegacyPptShape> sourceShapes,
            IReadOnlyDictionary<uint, uint> projectedShapeIds,
            LegacyPptSoundProjectionContext soundContext) {
            var animations = new List<PowerPointClassicAnimation>();
            foreach (LegacyPptShape shape in EnumerateLegacyShapes(sourceShapes)) {
                LegacyPptAnimation? animation = shape.Animation;
                if (animation == null || animation.HasSoundOverride
                    || !projectedShapeIds.TryGetValue(shape.ShapeId,
                        out uint projectedShapeId)) continue;
                string? relationshipId = null;
                string? soundName = null;
                if (animation.PlaysSound
                    && soundContext.TryProject(slide.SlidePart,
                        animation.SoundIdReference,
                        out LegacyPptSound? sound, out relationshipId)) {
                    soundName = sound?.Name;
                }
                animations.Add(new PowerPointClassicAnimation(projectedShapeId,
                    (PowerPointClassicAnimationEffect)(byte)animation.Effect,
                    animation.EffectDirection,
                    (PowerPointClassicAnimationBuildType)(byte)animation.BuildType,
                    animation.Automatic, Math.Max(0, animation.DelayMilliseconds),
                    animation.Order,
                    animation.Reverse, animation.AnimateBackground,
                    (PowerPointClassicAnimationAfterEffect)(byte)animation.AfterEffect,
                    (PowerPointClassicTextBuild)(byte)animation.TextBuildSubEffect,
                    animation.RawDimColor, animation.PlaysSound,
                    animation.StopsSound, relationshipId, soundName));
            }
            if (animations.Count > 0) slide.SetClassicAnimations(animations);
        }

        private static IEnumerable<LegacyPptShape> EnumerateLegacyShapes(
            IEnumerable<LegacyPptShape> roots) {
            foreach (LegacyPptShape shape in roots) {
                yield return shape;
                foreach (LegacyPptShape child in EnumerateLegacyShapes(shape.Children)) {
                    yield return child;
                }
            }
        }
    }
}
