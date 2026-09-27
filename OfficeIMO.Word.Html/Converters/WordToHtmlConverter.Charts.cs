using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static IElement? CreateChartImage(IDocument owner, WordChart chart,
            WordToHtmlOptions options, ref long embeddedImageBytes) {
            if (!chart.TryGetOfficeSnapshot(out var snapshot)) {
                AddExportDiagnostic(options, "WordChartOmitted",
                    "A Word chart was omitted because its complete cached data and appearance cannot be projected.",
                    OfficeConversionLossKind.Omission);
                return null;
            }

            const string prefix = "data:image/svg+xml;base64,";
            long remainingImages = options.MaxTotalEmbeddedImageBytes - embeddedImageBytes;
            long outputBytes = Math.Max(0, (GetRemainingOutputCharacters(owner) - prefix.Length) / 4L) * 3L;
            long maximumBytes = Math.Min(options.MaxEmbeddedImageBytes, Math.Min(remainingImages, outputBytes));
            string limitCode = maximumBytes == outputBytes ? "WordHtmlOutputLimitExceeded" :
                maximumBytes == remainingImages ? "WordImageTotalSizeLimitExceeded" : "WordImageSizeLimitExceeded";
            if (maximumBytes < 1) {
                long limit = limitCode == "WordHtmlOutputLimitExceeded" ? options.MaxOutputCharacters :
                    limitCode == "WordImageTotalSizeLimitExceeded" ? options.MaxTotalEmbeddedImageBytes : options.MaxEmbeddedImageBytes;
                long actual = limitCode == "WordHtmlOutputLimitExceeded"
                    ? SaturatingAdd(options.MaxOutputCharacters - GetRemainingOutputCharacters(owner), prefix.Length + 4L)
                    : limitCode == "WordImageTotalSizeLimitExceeded" ? SaturatingAdd(embeddedImageBytes, 1L) : 1L;
                ThrowExportLimitExceeded(options, limitCode, "A rendered chart cannot fit within the configured HTML export limits.", "WordChart", actual, limit);
            }
            byte[] bytes;
            try {
                var rendering = OfficeChartDrawingRenderer.RenderWithQuality(snapshot, useMinimumCanvas: false);
                if (rendering.QualityReport.HasIssues) {
                    AddExportDiagnostic(options, "WordChartRenderingApproximation",
                        "A Word chart was rendered with shared drawing quality warnings.", OfficeConversionLossKind.Approximation);
                }
                bytes = OfficeDrawingSvgExporter.ToSvgBytes(rendering.Drawing, 1D, OfficeSvgSizeUnit.Point,
                    null, null, maximumBytes, System.Threading.CancellationToken.None);
            } catch (OfficeImageExportBatchLimitException ex) {
                long actual = ex.Actual;
                long limit = options.MaxEmbeddedImageBytes;
                if (limitCode == "WordImageTotalSizeLimitExceeded") {
                    actual = SaturatingAdd(embeddedImageBytes, actual);
                    limit = options.MaxTotalEmbeddedImageBytes;
                } else if (limitCode == "WordHtmlOutputLimitExceeded") {
                    long encoded = actual > (long.MaxValue / 4L) * 3L ? long.MaxValue : ((actual + 2L) / 3L) * 4L;
                    actual = SaturatingAdd(options.MaxOutputCharacters - GetRemainingOutputCharacters(owner), SaturatingAdd(prefix.Length, encoded));
                    limit = options.MaxOutputCharacters;
                }
                ThrowExportLimitExceeded(options, limitCode, "A rendered chart exceeds the configured HTML export limits.",
                    "WordChart", actual, limit);
                throw;
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException || ex is NotSupportedException) {
                AddExportDiagnostic(options, "WordChartOmitted", "A Word chart could not be rendered as an HTML image.",
                    OfficeConversionLossKind.Omission);
                return null;
            }

            if (bytes.LongLength > options.MaxEmbeddedImageBytes)
                ThrowExportLimitExceeded(options, "WordImageSizeLimitExceeded", "A rendered chart exceeds the per-image HTML export limit.",
                    "WordChart", bytes.LongLength, options.MaxEmbeddedImageBytes);
            if (bytes.LongLength > options.MaxTotalEmbeddedImageBytes - embeddedImageBytes)
                ThrowExportLimitExceeded(options, "WordImageTotalSizeLimitExceeded", "Rendered charts and embedded images exceed the aggregate HTML export limit.",
                    "WordChart", SaturatingAdd(embeddedImageBytes, bytes.LongLength), options.MaxTotalEmbeddedImageBytes);

            long characters = prefix.Length + ((bytes.LongLength + 2L) / 3L) * 4L;
            ReserveOutputCharacters(owner, characters, "A rendered chart exceeds the HTML output-character limit.", "WordChart:src");
            var image = CreateOutputElement(owner, "img");
            SetOutputAttributeAfterValueReservation(owner, image, "src", prefix + System.Convert.ToBase64String(bytes), "WordChart:src");
            string alternativeText = !string.IsNullOrWhiteSpace(chart.AltText) ? chart.AltText! :
                string.IsNullOrWhiteSpace(snapshot.Title) ? "Word chart" : snapshot.Title!;
            SetOutputAttribute(image, "alt", alternativeText, "WordChart:alt");
            SetOutputAttribute(image, "width", Math.Round(snapshot.WidthPoints * 96D / 72D).ToString(CultureInfo.InvariantCulture), "WordChart:width");
            SetOutputAttribute(image, "height", Math.Round(snapshot.HeightPoints * 96D / 72D).ToString(CultureInfo.InvariantCulture), "WordChart:height");
            embeddedImageBytes += bytes.LongLength;
            AddExportDiagnostic(options, "WordChartRenderedAsImage",
                "A Word chart is represented by a static SVG image; its editable data and native layout are not preserved in HTML.",
                OfficeConversionLossKind.Approximation);
            return image;
        }
    }
}
