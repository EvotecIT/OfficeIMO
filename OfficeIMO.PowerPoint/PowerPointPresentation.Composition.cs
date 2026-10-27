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
            if (options.ValidatePlan) {
                IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics = resolvedPlan.ValidateSlides();
                if (diagnostics.Any(diagnostic =>
                        diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error)) {
                    throw new PowerPointDeckPlanValidationException(diagnostics);
                }
            }
            PowerPointDeckDesign design = options.ResolveDesign(resolvedPlan);
            var composer = new PowerPointDeckComposer(this, design, options.ApplyTheme, options.TemplateLayouts);
            int diagnosticStart = LayoutDeck.Diagnostics.Count;
            IReadOnlyList<PowerPointSlide> slides = composer.AddSlides(resolvedPlan, validate: false);
            PowerPointDeckPreflightReport preflight = WithLayoutDiagnostics(InspectPreflight(options.Preflight), diagnosticStart);
            return new PowerPointCompositionResult(resolvedPlan, design, slides.ToList(), preflight);
        }

        /// <summary>Adds what the layout system reported while composing (splits, overflow, ignored options) to the preflight report.</summary>
        private PowerPointDeckPreflightReport WithLayoutDiagnostics(PowerPointDeckPreflightReport preflight, int start) {
            PowerPointLayoutDeck deck = LayoutDeck;
            if (deck.Diagnostics.Count <= start) return preflight;
            var findings = new List<PowerPointDeckPreflightFinding>(preflight.Findings);
            for (int i = start; i < deck.Diagnostics.Count; i++) {
                OfficeDiagnostic diagnostic = deck.Diagnostics[i];
                PowerPointDeckPreflightSeverity severity = diagnostic.Severity switch {
                    OfficeDiagnosticSeverity.Error => PowerPointDeckPreflightSeverity.Error,
                    OfficeDiagnosticSeverity.Warning => PowerPointDeckPreflightSeverity.Warning,
                    _ => PowerPointDeckPreflightSeverity.Info
                };
                findings.Add(new PowerPointDeckPreflightFinding(severity, "Layout." + diagnostic.Code, diagnostic.Message,
                    Math.Min(deck.DiagnosticSlides[i], Math.Max(0, Slides.Count - 1))));
            }

            return new PowerPointDeckPreflightReport(preflight.SlideCount, findings);
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
