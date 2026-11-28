using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacyTransition(PowerPointSlide slide,
            LegacyPptTransition? source,
            LegacyPptSoundProjectionContext soundContext) {
            if (source == null) return;
            PowerPointSlideTransition? transition =
                LegacyPptTransitionMapping.ToSlideTransition(source);
            if (!transition.HasValue) return;
            slide.Transition = transition.Value;
            slide.TransitionSpeed = source.Speed switch {
                0 => PowerPointSlideTransitionSpeed.Slow,
                2 => PowerPointSlideTransitionSpeed.Fast,
                _ => PowerPointSlideTransitionSpeed.Medium
            };
            slide.TransitionAdvanceOnClick = source.ManualAdvance;
            slide.TransitionAdvanceAfterSeconds = source.AutoAdvance
                ? source.SlideTimeMilliseconds / 1000.0
                : null;
            PowerPointPartRef slidePart = slide.PartRef;
            OfficeXmlNode target = slidePart.Root.Child(LegacyPptXml.P, "transition");
            if (target.IsNone) return;
            if (source.PlaySound
                       && soundContext.TryProject(slidePart, source.SoundId,
                           out LegacyPptSound? sound, out string? relationshipId)) {
                target.RemoveChildren(LegacyPptXml.P, "sndAc");
                target.AddP("sndAc").AddP("stSnd")
                    .Set("loop", source.LoopSound)
                    .AddP("snd")
                    .SetRelationship("embed", relationshipId)
                    .Set("name", sound!.Name)
                    .Set("builtIn", sound.BuiltInId.HasValue);
                slidePart.MarkEdited();
            } else if (source.StopSound) {
                target.RemoveChildren(LegacyPptXml.P, "sndAc");
                target.AddP("sndAc").AddP("endSnd");
                slidePart.MarkEdited();
            }
        }
    }
}
