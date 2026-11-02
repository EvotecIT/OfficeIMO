using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ApplyCachedImageSize(WordImage image, WordImage cached, double? width, double? height) {
            if (!_unscaledImageSizes.TryGetValue(cached, out var original)) {
                original = ((cached.Width?.ToPixels()) ?? 0D, (cached.Height?.ToPixels()) ?? 0D);
            }
            ApplyImageSize(image, original, width, height);
        }

        private static void ApplyImageSize(WordImage image, (double Width, double Height) original,
            double? width, double? height) {
            if (original.Width <= 0D || original.Height <= 0D) return;

            if (width.HasValue && height.HasValue) {
                image.Width = OfficeIMO.OfficeLength.Pixels(width.Value);
                image.Height = OfficeIMO.OfficeLength.Pixels(height.Value);
            } else if (width.HasValue) {
                image.Width = OfficeIMO.OfficeLength.Pixels(width.Value);
                image.Height = OfficeIMO.OfficeLength.Pixels(original.Height * width.Value / original.Width);
            } else if (height.HasValue) {
                image.Width = OfficeIMO.OfficeLength.Pixels(original.Width * height.Value / original.Height);
                image.Height = OfficeIMO.OfficeLength.Pixels(height.Value);
            } else {
                image.Width = OfficeIMO.OfficeLength.Pixels(original.Width);
                image.Height = OfficeIMO.OfficeLength.Pixels(original.Height);
            }
        }

        /// <summary>
        /// Bounds an image to its supported CSS maximum width, or a dimensionless
        /// image to the Word content width, while preserving its aspect ratio.
        /// </summary>
        private static void FitImageToWidthConstraints(
            WordImage image, double? contentWidth, double? maximumWidth, WordHtmlImportOptions options,
            string source, string alternativeText) {
            double? limit = contentWidth;
            if (maximumWidth is > 0D) {
                limit = limit.HasValue ? Math.Min(limit.Value, maximumWidth.Value) : maximumWidth;
            }
            if (limit is not > 0D || (image.Width?.ToPixels()) is not double width ||
                (image.Height?.ToPixels()) is not double height || width <= limit.Value || height <= 0D) {
                return;
            }

            double scale = limit.Value / width;
            image.Width = OfficeIMO.OfficeLength.Pixels(limit.Value);
            image.Height = OfficeIMO.OfficeLength.Pixels(height * scale);
            string diagnosticSource = source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                ? string.IsNullOrWhiteSpace(alternativeText) ? "data:image" : alternativeText
                : source;
            AddDiagnostic(options, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Image width exceeded its supported CSS maximum or the Word content width and was scaled proportionally to fit.",
                diagnosticSource, lossKind: OfficeConversionLossKind.Approximation);
        }
    }
}
