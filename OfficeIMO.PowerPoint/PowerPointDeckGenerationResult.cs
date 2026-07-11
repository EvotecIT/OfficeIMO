using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Rendered semantic slides paired with the shared machine-readable deck report.</summary>
    public sealed class PowerPointDeckGenerationResult {
        internal PowerPointDeckGenerationResult(IList<PowerPointSlide> slides,
            PowerPointDeckPreflightReport report) {
            Slides = new ReadOnlyCollection<PowerPointSlide>(
                new List<PowerPointSlide>(slides ?? throw new ArgumentNullException(nameof(slides))));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Rendered slides in generation order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides { get; }

        /// <summary>Shared deterministic preflight report for the resulting deck.</summary>
        public PowerPointDeckPreflightReport Report { get; }
    }
}
