using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyTransition(PowerPointSlide slide,
            LegacyPptTransition? source) {
            if (source == null) return;
            SlideTransition? transition =
                LegacyPptTransitionMapping.ToSlideTransition(source);
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

    }
}
