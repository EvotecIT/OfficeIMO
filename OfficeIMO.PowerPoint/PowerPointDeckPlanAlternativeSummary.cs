using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight description of how one deck plan resolves under one generated design alternative.
    /// </summary>
    public sealed class PowerPointDeckPlanAlternativeSummary {
        internal PowerPointDeckPlanAlternativeSummary(int index, PowerPointDesignVariety variety,
            PowerPointDeckDesignSummary design,
            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics,
            int contentFitScore,
            IReadOnlyList<string> contentFitReasons) {
            Index = index;
            Variety = variety;
            Design = design;
            Slides = slides;
            Diagnostics = diagnostics;
            ContentFitScore = contentFitScore;
            ContentFitReasons = contentFitReasons;
        }

        /// <summary>
        ///     Zero-based alternative index.
        /// </summary>
        public int Index { get; }

        /// <summary>
        ///     Variety level used when resolving design alternatives for this preview.
        /// </summary>
        public PowerPointDesignVariety Variety { get; }

        /// <summary>
        ///     Design alternative summary.
        /// </summary>
        public PowerPointDeckDesignSummary Design { get; }

        /// <summary>
        ///     Slide render summaries resolved under this design alternative.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> Slides { get; }

        /// <summary>
        ///     Plan diagnostics shared by every alternative.
        /// </summary>
        public IReadOnlyList<PowerPointDeckPlanDiagnostic> Diagnostics { get; }

        /// <summary>
        ///     Lightweight score describing how well this design alternative fits the planned slide content.
        /// </summary>
        public int ContentFitScore { get; }

        /// <summary>
        ///     Short explanations for why this design alternative fits the planned slide content.
        /// </summary>
        public IReadOnlyList<string> ContentFitReasons { get; }

        /// <summary>
        ///     Whether this alternative has at least one positive content-fit signal.
        /// </summary>
        public bool MatchesContent => ContentFitScore > 0;

        /// <summary>
        ///     Whether the plan has any diagnostics that would prevent semantic rendering.
        /// </summary>
        public bool HasErrors => Diagnostics.Any(diagnostic =>
            diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error);

        /// <summary>
        ///     Whether the plan has any warning diagnostics.
        /// </summary>
        public bool HasWarnings => Diagnostics.Any(diagnostic =>
            diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Warning);

        /// <inheritdoc />
        public override string ToString() {
            return Index + ": " + Design.DirectionName + " (" + Variety + ") - " + Slides.Count +
                   " slides, fit " + ContentFitScore;
        }
    }
}
