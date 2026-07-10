using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Options controlling dependency-free PowerPoint slide image export.
    /// </summary>
    public class PowerPointImageExportOptions : OfficeImageExportOptions {
        /// <summary>
        /// Gets or sets a value indicating whether resolved slide backgrounds should be included.
        /// </summary>
        public bool IncludeSlideBackground { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether slide content should be rendered when supported.
        /// </summary>
        public bool IncludeSlideContent { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether hidden slide shapes should be included.
        /// </summary>
        public bool IncludeHiddenShapes { get; set; }

        internal PowerPointImageExportOptions Clone() => new PowerPointImageExportOptions {
            Scale = Scale,
            BackgroundColor = BackgroundColor,
            IncludeSlideBackground = IncludeSlideBackground,
            IncludeSlideContent = IncludeSlideContent,
            IncludeHiddenShapes = IncludeHiddenShapes
        };
    }

    /// <summary>
    /// Options controlling dependency-free PowerPoint presentation image export.
    /// </summary>
    public sealed class PowerPointPresentationImageExportOptions : PowerPointImageExportOptions {
        /// <summary>
        /// Gets or sets a value indicating whether hidden slides should be included in presentation-level exports.
        /// </summary>
        public bool IncludeHiddenSlides { get; set; }

        /// <summary>
        /// Gets or sets the 1-based slide numbers to export. When null or empty, all eligible slides are exported.
        /// </summary>
        public IReadOnlyList<int>? SlideNumbers { get; set; }

        internal PowerPointPresentationImageExportOptions ClonePresentation() => new PowerPointPresentationImageExportOptions {
            Scale = Scale,
            BackgroundColor = BackgroundColor,
            IncludeSlideBackground = IncludeSlideBackground,
            IncludeSlideContent = IncludeSlideContent,
            IncludeHiddenShapes = IncludeHiddenShapes,
            IncludeHiddenSlides = IncludeHiddenSlides,
            SlideNumbers = SlideNumbers?.ToArray()
        };
    }
}
