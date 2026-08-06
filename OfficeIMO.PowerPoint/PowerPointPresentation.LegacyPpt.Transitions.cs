using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
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
            DocumentFormat.OpenXml.Presentation.Transition? target =
                slide.SlidePart.Slide?.Transition;
            if (target == null) return;
            if (source.PlaySound
                       && soundContext.TryProject(slide.SlidePart, source.SoundId,
                           out LegacyPptSound? sound, out string? relationshipId)) {
                target.RemoveAllChildren<DocumentFormat.OpenXml.Presentation.SoundAction>();
                var embeddedSound = new DocumentFormat.OpenXml.Presentation.Sound {
                    Embed = relationshipId,
                    Name = sound!.Name,
                    BuiltIn = sound.BuiltInId.HasValue
                };
                target.RemoveAllChildren<DocumentFormat.OpenXml.Presentation.SoundAction>();
                target.Append(new DocumentFormat.OpenXml.Presentation.SoundAction(
                    new DocumentFormat.OpenXml.Presentation.StartSoundAction(
                        embeddedSound) { Loop = source.LoopSound }));
            } else if (source.StopSound) {
                target.RemoveAllChildren<DocumentFormat.OpenXml.Presentation.SoundAction>();
                target.Append(new DocumentFormat.OpenXml.Presentation.SoundAction(
                    new DocumentFormat.OpenXml.Presentation.EndSoundAction()));
            }
        }

    }
}
