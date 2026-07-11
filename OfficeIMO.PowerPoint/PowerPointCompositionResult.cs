using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint {
    /// <summary>Complete result of composing a semantic deck plan into a presentation.</summary>
    public sealed class PowerPointCompositionResult {
        internal PowerPointCompositionResult(PowerPointDeckPlan plan, PowerPointDeckDesign design,
            IList<PowerPointSlide> slides, PowerPointDeckPreflightReport preflight) {
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Design = design ?? throw new ArgumentNullException(nameof(design));
            Slides = new ReadOnlyCollection<PowerPointSlide>(
                new List<PowerPointSlide>(slides ?? throw new ArgumentNullException(nameof(slides))));
            Preflight = preflight ?? throw new ArgumentNullException(nameof(preflight));
        }

        /// <summary>Resolved plan, including continuation slides when requested.</summary>
        public PowerPointDeckPlan Plan { get; }

        /// <summary>Resolved design used by the shared composition engine.</summary>
        public PowerPointDeckDesign Design { get; }

        /// <summary>Rendered slides in generation order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides { get; }

        /// <summary>Deterministic preflight report for the resulting presentation.</summary>
        public PowerPointDeckPreflightReport Preflight { get; }
    }
}
