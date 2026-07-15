using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyTransition(PowerPointSlide slide,
            LegacyPptTransition? source) {
            if (source == null) return;
            SlideTransition? transition = MapLegacyTransition(source);
            if (!transition.HasValue) return;
            slide.Transition = transition.Value;
            slide.TransitionSpeed = source.Speed switch {
                0 => SlideTransitionSpeed.Slow,
                2 => SlideTransitionSpeed.Fast,
                _ => SlideTransitionSpeed.Medium
            };
            slide.TransitionAdvanceOnClick = source.ManualAdvance;
            slide.TransitionAdvanceAfterSeconds = source.AutoAdvance
                ? source.SlideTimeMilliseconds / 1000.0
                : null;
        }

        private static SlideTransition? MapLegacyTransition(
            LegacyPptTransition source) => source.Effect switch {
                LegacyPptTransitionEffect.Cut => SlideTransition.Cut,
                LegacyPptTransitionEffect.Fade => SlideTransition.Fade,
                LegacyPptTransitionEffect.Wipe => SlideTransition.Wipe,
                LegacyPptTransitionEffect.Blinds => source.EffectDirection == 0
                    ? SlideTransition.BlindsVertical
                    : SlideTransition.BlindsHorizontal,
                LegacyPptTransitionEffect.Comb => source.EffectDirection == 0
                    ? SlideTransition.CombHorizontal
                    : SlideTransition.CombVertical,
                LegacyPptTransitionEffect.Push => source.EffectDirection switch {
                    1 => SlideTransition.PushUp,
                    2 => SlideTransition.PushRight,
                    3 => SlideTransition.PushDown,
                    _ => SlideTransition.PushLeft
                },
                _ => null
            };
    }
}
