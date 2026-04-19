using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Lightweight description of how one deck plan resolves under one generated design alternative.
    /// </summary>
    public sealed class PowerPointDeckPlanAlternativeSummary {
        internal PowerPointDeckPlanAlternativeSummary(int index, PowerPointDeckDesignSummary design,
            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> slides,
            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics) {
            Index = index;
            Design = design;
            Slides = slides;
            Diagnostics = diagnostics;
        }

        /// <summary>
        ///     Zero-based alternative index.
        /// </summary>
        public int Index { get; }

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
            return Index + ": " + Design.DirectionName + " - " + Slides.Count + " slides";
        }
    }
}
