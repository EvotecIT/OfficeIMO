using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Composes a semantic deck plan into this presentation through the single shared layout engine.
        /// </summary>
        /// <param name="plan">Semantic slide intent to render.</param>
        /// <param name="options">Design, continuation, template-layout, validation, and preflight policy.</param>
        /// <returns>The resolved plan, design, rendered slides, and preflight report.</returns>
        public PowerPointCompositionResult Compose(PowerPointDeckPlan plan, PowerPointCompositionOptions options) {
            ThrowIfDisposed();
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (options == null) throw new ArgumentNullException(nameof(options));

            PowerPointDeckPlan resolvedPlan = options.ExpandContinuations
                ? plan.WithContinuations(options.Continuation)
                : plan;
            PowerPointDeckDesign design = options.ResolveDesign(resolvedPlan);
            var composer = new PowerPointDeckComposer(this, design, options.ApplyTheme, options.TemplateLayouts);
            IReadOnlyList<PowerPointSlide> slides = composer.AddSlides(resolvedPlan, options.ValidatePlan);
            PowerPointDeckPreflightReport preflight = InspectPreflight(options.Preflight);
            return new PowerPointCompositionResult(resolvedPlan, design, slides.ToList(), preflight);
        }

        /// <summary>
        ///     Describes how a semantic plan will resolve without changing this presentation.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> PreviewComposition(
            PowerPointDeckPlan plan, PowerPointCompositionOptions options) {
            ThrowIfDisposed();
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (options == null) throw new ArgumentNullException(nameof(options));

            PowerPointDeckPlan resolvedPlan = options.ExpandContinuations
                ? plan.WithContinuations(options.Continuation)
                : plan;
            PowerPointDeckDesign design = options.ResolveDesign(resolvedPlan);
            return resolvedPlan.DescribeSlides(design, Slides.Count);
        }
    }
}
