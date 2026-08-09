using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Fluent image-export builder for a PowerPoint slide.
    /// </summary>
    public sealed class PowerPointSlideImageExportBuilder : OfficeImageExportBuilder<PowerPointSlideImageExportBuilder, PowerPointImageExportOptions> {
        internal PowerPointSlideImageExportBuilder(PowerPointSlide slide, PowerPointImageExportOptions? options = null)
            : base(
                options?.Clone() ?? new PowerPointImageExportOptions(),
                (format, effective, cancellationToken) => slide.ExportImage(format, effective, cancellationToken)) {
        }

        /// <summary>Includes or excludes the resolved slide background.</summary>
        public PowerPointSlideImageExportBuilder IncludeBackground(bool include = true) {
            Options.IncludeSlideBackground = include;
            return this;
        }

        /// <summary>Excludes the resolved slide background.</summary>
        public PowerPointSlideImageExportBuilder WithoutBackground() => IncludeBackground(false);

        /// <summary>Includes or excludes slide content.</summary>
        public PowerPointSlideImageExportBuilder IncludeContent(bool include = true) {
            Options.IncludeSlideContent = include;
            return this;
        }

        /// <summary>Excludes slide content.</summary>
        public PowerPointSlideImageExportBuilder WithoutContent() => IncludeContent(false);

        /// <summary>Includes or excludes pictures.</summary>
        public PowerPointSlideImageExportBuilder IncludePictures(bool include = true) { Options.IncludePictures = include; return this; }

        /// <summary>Includes or excludes auto shapes.</summary>
        public PowerPointSlideImageExportBuilder IncludeAutoShapes(bool include = true) { Options.IncludeAutoShapes = include; return this; }

        /// <summary>Includes or excludes SmartArt.</summary>
        public PowerPointSlideImageExportBuilder IncludeSmartArt(bool include = true) { Options.IncludeSmartArt = include; return this; }

        /// <summary>Includes or excludes text boxes.</summary>
        public PowerPointSlideImageExportBuilder IncludeTextBoxes(bool include = true) { Options.IncludeTextBoxes = include; return this; }

        /// <summary>Includes or excludes tables.</summary>
        public PowerPointSlideImageExportBuilder IncludeTables(bool include = true) { Options.IncludeTables = include; return this; }

        /// <summary>Includes or excludes charts.</summary>
        public PowerPointSlideImageExportBuilder IncludeCharts(bool include = true) { Options.IncludeCharts = include; return this; }

        /// <summary>Includes or excludes shapes marked hidden.</summary>
        public PowerPointSlideImageExportBuilder IncludeHiddenShapes(bool include = true) { Options.IncludeHiddenShapes = include; return this; }

        /// <summary>Sets the maximum nested group depth accepted during rendering.</summary>
        public PowerPointSlideImageExportBuilder WithMaximumGroupDepth(int maximumDepth) {
            Options.MaxGroupShapeDepth = maximumDepth;
            return this;
        }

        /// <summary>Sets per-image and aggregate embedded-image read limits.</summary>
        public PowerPointSlideImageExportBuilder WithEmbeddedImageLimits(int maximumImageBytes, long maximumTotalImageBytes) {
            if (maximumImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumImageBytes));
            if (maximumTotalImageBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumTotalImageBytes));
            Options.MaximumEmbeddedImageBytes = maximumImageBytes;
            Options.MaximumTotalEmbeddedImageBytes = maximumTotalImageBytes;
            return this;
        }
    }

    /// <summary>
    /// Fluent image-export builder for all selected slides in a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointPresentationImageExportBuilder : OfficeImageExportBatchBuilder<PowerPointPresentationImageExportBuilder, PowerPointPresentationImageExportOptions> {
        internal PowerPointPresentationImageExportBuilder(PowerPointPresentation presentation, PowerPointPresentationImageExportOptions? options = null)
            : base(
                options?.ClonePresentation() ?? new PowerPointPresentationImageExportOptions(),
                presentation.ExportImages,
                (format, effective, consumer, cancellationToken) =>
                    presentation.ExportImages(format, consumer, effective, cancellationToken)) {
        }

        /// <summary>Includes or excludes resolved slide backgrounds.</summary>
        public PowerPointPresentationImageExportBuilder IncludeBackground(bool include = true) {
            Options.IncludeSlideBackground = include;
            return this;
        }

        /// <summary>Excludes resolved slide backgrounds.</summary>
        public PowerPointPresentationImageExportBuilder WithoutBackground() => IncludeBackground(false);

        /// <summary>Includes or excludes slide content.</summary>
        public PowerPointPresentationImageExportBuilder IncludeContent(bool include = true) {
            Options.IncludeSlideContent = include;
            return this;
        }

        /// <summary>Excludes slide content.</summary>
        public PowerPointPresentationImageExportBuilder WithoutContent() => IncludeContent(false);

        /// <summary>Includes or excludes pictures.</summary>
        public PowerPointPresentationImageExportBuilder IncludePictures(bool include = true) { Options.IncludePictures = include; return this; }

        /// <summary>Includes or excludes auto shapes.</summary>
        public PowerPointPresentationImageExportBuilder IncludeAutoShapes(bool include = true) { Options.IncludeAutoShapes = include; return this; }

        /// <summary>Includes or excludes SmartArt.</summary>
        public PowerPointPresentationImageExportBuilder IncludeSmartArt(bool include = true) { Options.IncludeSmartArt = include; return this; }

        /// <summary>Includes or excludes text boxes.</summary>
        public PowerPointPresentationImageExportBuilder IncludeTextBoxes(bool include = true) { Options.IncludeTextBoxes = include; return this; }

        /// <summary>Includes or excludes tables.</summary>
        public PowerPointPresentationImageExportBuilder IncludeTables(bool include = true) { Options.IncludeTables = include; return this; }

        /// <summary>Includes or excludes charts.</summary>
        public PowerPointPresentationImageExportBuilder IncludeCharts(bool include = true) { Options.IncludeCharts = include; return this; }

        /// <summary>Includes or excludes shapes marked hidden.</summary>
        public PowerPointPresentationImageExportBuilder IncludeHiddenShapes(bool include = true) { Options.IncludeHiddenShapes = include; return this; }

        /// <summary>Sets the maximum nested group depth accepted during rendering.</summary>
        public PowerPointPresentationImageExportBuilder WithMaximumGroupDepth(int maximumDepth) {
            Options.MaxGroupShapeDepth = maximumDepth;
            return this;
        }

        /// <summary>Sets per-image and aggregate embedded-image read limits.</summary>
        public PowerPointPresentationImageExportBuilder WithEmbeddedImageLimits(int maximumImageBytes, long maximumTotalImageBytes) {
            if (maximumImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumImageBytes));
            if (maximumTotalImageBytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumTotalImageBytes));
            Options.MaximumEmbeddedImageBytes = maximumImageBytes;
            Options.MaximumTotalEmbeddedImageBytes = maximumTotalImageBytes;
            return this;
        }

        /// <summary>Includes or excludes hidden slides.</summary>
        public PowerPointPresentationImageExportBuilder IncludeHiddenSlides(bool include = true) {
            Options.IncludeHiddenSlides = include;
            return this;
        }

        /// <summary>Selects specific 1-based slide numbers for presentation-level export.</summary>
        public PowerPointPresentationImageExportBuilder ForSlides(params int[] slideNumbers) => ForSlides((IEnumerable<int>)slideNumbers);

        /// <summary>Selects specific 1-based slide numbers for presentation-level export.</summary>
        public PowerPointPresentationImageExportBuilder ForSlides(IEnumerable<int> slideNumbers) {
            if (slideNumbers == null) {
                throw new ArgumentNullException(nameof(slideNumbers));
            }

            int[] selected = slideNumbers.ToArray();
            if (selected.Length == 0) {
                throw new ArgumentException("At least one slide number must be specified.", nameof(slideNumbers));
            }

            if (selected.Any(slideNumber => slideNumber <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(slideNumbers), "Slide numbers must be 1-based positive values.");
            }

            Options.SlideNumbers = selected;
            return this;
        }

        /// <summary>Selects a 1-based inclusive slide range for presentation-level export.</summary>
        public PowerPointPresentationImageExportBuilder ForSlideRange(int firstSlideNumber, int lastSlideNumber) {
            if (firstSlideNumber <= 0) {
                throw new ArgumentOutOfRangeException(nameof(firstSlideNumber), "Slide numbers must be 1-based positive values.");
            }

            if (lastSlideNumber < firstSlideNumber) {
                throw new ArgumentOutOfRangeException(nameof(lastSlideNumber), "The last slide number must be greater than or equal to the first slide number.");
            }

            return ForSlides(Enumerable.Range(firstSlideNumber, lastSlideNumber - firstSlideNumber + 1));
        }
    }

    public partial class PowerPointSlide {
        /// <summary>
        /// Starts a fluent image export for this slide.
        /// </summary>
        public PowerPointSlideImageExportBuilder ToImage() => new PowerPointSlideImageExportBuilder(this);

        /// <summary>Starts a fluent image export using a cloned options snapshot.</summary>
        public PowerPointSlideImageExportBuilder ToImage(PowerPointImageExportOptions options) =>
            new PowerPointSlideImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>
        /// Starts a fluent image export for selected slides in this presentation.
        /// </summary>
        public PowerPointPresentationImageExportBuilder ToImages() => new PowerPointPresentationImageExportBuilder(this);

        /// <summary>Starts a fluent batch export using a cloned options snapshot.</summary>
        public PowerPointPresentationImageExportBuilder ToImages(PowerPointPresentationImageExportOptions options) =>
            new PowerPointPresentationImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }
}
