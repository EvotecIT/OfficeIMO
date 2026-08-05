using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Options controlling dependency-free PowerPoint slide image export.
    /// </summary>
    public class PowerPointImageExportOptions : OfficeImageExportOptions {
        /// <summary>Default maximum bytes read from one embedded image during export.</summary>
        public const int DefaultMaximumEmbeddedImageBytes = 64 * 1024 * 1024;

        /// <summary>Default aggregate embedded-image bytes retained by one slide export.</summary>
        public const long DefaultMaximumTotalEmbeddedImageBytes = 256L * 1024 * 1024;

        /// <inheritdoc />
        public override double LogicalUnitsPerInch => 72D;

        /// <summary>
        /// Gets or sets a value indicating whether resolved slide backgrounds should be included.
        /// </summary>
        public bool IncludeSlideBackground { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether slide content should be rendered when supported.
        /// </summary>
        public bool IncludeSlideContent { get; set; } = true;

        /// <summary>Gets or sets whether pictures and media should be rendered.</summary>
        public bool IncludePictures { get; set; } = true;

        /// <summary>Gets or sets whether auto-shapes and connectors should be rendered.</summary>
        public bool IncludeAutoShapes { get; set; } = true;

        /// <summary>Gets or sets whether supported SmartArt diagrams should be rendered.</summary>
        public bool IncludeSmartArt { get; set; } = true;

        /// <summary>Gets or sets whether text boxes should be rendered.</summary>
        public bool IncludeTextBoxes { get; set; } = true;

        /// <summary>Gets or sets whether tables should be rendered.</summary>
        public bool IncludeTables { get; set; } = true;

        /// <summary>Gets or sets whether charts should be rendered.</summary>
        public bool IncludeCharts { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether hidden slide shapes should be included.
        /// </summary>
        public bool IncludeHiddenShapes { get; set; }

        /// <summary>Maximum nested group-shape depth rendered into a shared visual snapshot. Zero skips group content; negative values disable the depth limit.</summary>
        public int MaxGroupShapeDepth { get; set; } = 32;

        /// <summary>Maximum bytes read from one embedded image during export.</summary>
        public int MaximumEmbeddedImageBytes { get; set; } = DefaultMaximumEmbeddedImageBytes;

        /// <summary>Maximum aggregate embedded-image bytes retained by one slide export.</summary>
        public long MaximumTotalEmbeddedImageBytes { get; set; } = DefaultMaximumTotalEmbeddedImageBytes;

        /// <summary>Creates an independent options snapshot.</summary>
        public PowerPointImageExportOptions Clone() =>
            CopyPowerPointOptionsTo(new PowerPointImageExportOptions());

        internal T CopyPowerPointOptionsTo<T>(T target) where T : PowerPointImageExportOptions {
            CopyImageExportOptionsTo(target);
            target.IncludeSlideBackground = IncludeSlideBackground;
            target.IncludeSlideContent = IncludeSlideContent;
            target.IncludePictures = IncludePictures;
            target.IncludeAutoShapes = IncludeAutoShapes;
            target.IncludeSmartArt = IncludeSmartArt;
            target.IncludeTextBoxes = IncludeTextBoxes;
            target.IncludeTables = IncludeTables;
            target.IncludeCharts = IncludeCharts;
            target.IncludeHiddenShapes = IncludeHiddenShapes;
            target.MaxGroupShapeDepth = MaxGroupShapeDepth;
            target.MaximumEmbeddedImageBytes = MaximumEmbeddedImageBytes;
            target.MaximumTotalEmbeddedImageBytes = MaximumTotalEmbeddedImageBytes;
            return target;
        }

        internal void Validate() {
            ValidateImageExportOptions();
            if (MaximumEmbeddedImageBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaximumEmbeddedImageBytes));
            }
            if (MaximumTotalEmbeddedImageBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaximumTotalEmbeddedImageBytes));
            }
        }
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

        /// <summary>Creates an independent presentation options snapshot.</summary>
        public PowerPointPresentationImageExportOptions ClonePresentation() {
            PowerPointPresentationImageExportOptions clone =
                CopyPowerPointOptionsTo(new PowerPointPresentationImageExportOptions());
            clone.IncludeHiddenSlides = IncludeHiddenSlides;
            clone.SlideNumbers = SlideNumbers?.ToArray();
            return clone;
        }
    }
}
