using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        /// Exports presentation slides as supported raster formats or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, PowerPointPresentationImageExportOptions? options = null) {
            ThrowIfDisposed();
            PowerPointPresentationImageExportOptions resolved = NormalizePresentationImageExportOptions(options);
            PowerPointImageExportOptions slideOptions = CreateSlideImageExportOptions(resolved);
            HashSet<int>? selectedSlideNumbers = CreateSelectedSlideNumberSet(resolved, Slides.Count);
            var results = new List<OfficeImageExportResult>();

            for (int i = 0; i < Slides.Count; i++) {
                int slideNumber = i + 1;
                if (selectedSlideNumbers != null && !selectedSlideNumbers.Contains(slideNumber)) {
                    continue;
                }

                PowerPointSlide slide = Slides[i];
                if (!resolved.IncludeHiddenSlides && slide.Hidden) {
                    continue;
                }

                OfficeImageExportResult result = slide.ExportImage(format, slideOptions);
                results.Add(new OfficeImageExportResult(
                    result.Format,
                    result.Width,
                    result.Height,
                    result.Bytes,
                    "Slide " + slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    "PowerPoint slide " + slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    result.Diagnostics));
            }

            return results.AsReadOnly();
        }

        /// <summary>
        /// Saves visible presentation slides as PNG files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, PowerPointPresentationImageExportOptions? options = null) =>
            new PowerPointPresentationImageExportBuilder(this, options).AsPng().Save(folderPath);

        /// <summary>
        /// Saves visible presentation slides as image files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, OfficeImageExportFormat format, PowerPointPresentationImageExportOptions? options = null) {
            PowerPointPresentationImageExportBuilder builder = new PowerPointPresentationImageExportBuilder(this, options);
            builder.As(format);
            return builder.Save(folderPath);
        }

        /// <summary>
        /// Asynchronously saves visible presentation slides as PNG files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            PowerPointPresentationImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointPresentationImageExportBuilder(this, options).AsPng().SaveAsync(folderPath, cancellationToken);

        /// <summary>
        /// Asynchronously saves visible presentation slides as image files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            OfficeImageExportFormat format,
            PowerPointPresentationImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            PowerPointPresentationImageExportBuilder builder = new PowerPointPresentationImageExportBuilder(this, options);
            builder.As(format);
            return builder.SaveAsync(folderPath, cancellationToken);
        }

        private static PowerPointPresentationImageExportOptions NormalizePresentationImageExportOptions(PowerPointPresentationImageExportOptions? options) {
            PowerPointPresentationImageExportOptions resolved = options?.ClonePresentation() ?? new PowerPointPresentationImageExportOptions();
            resolved.Validate();
            if (resolved.SlideNumbers != null && resolved.SlideNumbers.Any(slideNumber => slideNumber <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Slide numbers must be 1-based positive values.");
            }

            return resolved;
        }

        private static HashSet<int>? CreateSelectedSlideNumberSet(PowerPointPresentationImageExportOptions options, int slideCount) {
            if (options.SlideNumbers == null || options.SlideNumbers.Count == 0) {
                return null;
            }

            int highestSlideNumber = options.SlideNumbers.Max();
            if (highestSlideNumber > slideCount) {
                throw new ArgumentOutOfRangeException(nameof(options), "Selected slide numbers must exist in the presentation.");
            }

            return new HashSet<int>(options.SlideNumbers);
        }

        private static PowerPointImageExportOptions CreateSlideImageExportOptions(
            PowerPointPresentationImageExportOptions options) => options.Clone();
    }
}
