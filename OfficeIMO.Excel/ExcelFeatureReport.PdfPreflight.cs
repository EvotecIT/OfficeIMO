using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static bool IsPdfSupportedChartType(ExcelChartType chartType) {
            switch (chartType) {
                case ExcelChartType.ColumnClustered:
                case ExcelChartType.Column3DClustered:
                case ExcelChartType.ColumnStacked:
                case ExcelChartType.Column3DStacked:
                case ExcelChartType.ColumnStacked100:
                case ExcelChartType.Column3DStacked100:
                case ExcelChartType.BarClustered:
                case ExcelChartType.Bar3DClustered:
                case ExcelChartType.BarStacked:
                case ExcelChartType.Bar3DStacked:
                case ExcelChartType.BarStacked100:
                case ExcelChartType.Bar3DStacked100:
                case ExcelChartType.Line:
                case ExcelChartType.Line3D:
                case ExcelChartType.LineStacked:
                case ExcelChartType.LineStacked100:
                case ExcelChartType.Area:
                case ExcelChartType.Area3D:
                case ExcelChartType.AreaStacked:
                case ExcelChartType.Area3DStacked:
                case ExcelChartType.AreaStacked100:
                case ExcelChartType.Area3DStacked100:
                case ExcelChartType.Scatter:
                case ExcelChartType.Radar:
                case ExcelChartType.Pie:
                case ExcelChartType.Pie3D:
                case ExcelChartType.PieOfPie:
                case ExcelChartType.BarOfPie:
                case ExcelChartType.Doughnut:
                    return true;
                default:
                    return false;
            }
        }

        private static bool HasRenderablePdfChartData(ExcelChartSnapshot snapshot) {
            return snapshot.Data.Categories.Count > 0 && snapshot.Data.Series.Count > 0;
        }

        private static bool HasMixedPdfChartTypes(ExcelChartSnapshot snapshot) {
            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                if (series.ChartType.HasValue && series.ChartType.Value != snapshot.ChartType) {
                    return true;
                }
            }

            return false;
        }

        private static string GetChartDisplayName(ExcelChart chart) {
            if (!string.IsNullOrWhiteSpace(chart.Title)) {
                return chart.Title!;
            }

            return string.IsNullOrWhiteSpace(chart.Name) ? chart.ChartType.ToString() : chart.Name;
        }

        private static string GetChartDisplayName(ExcelChartSnapshot snapshot) {
            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            return string.IsNullOrWhiteSpace(snapshot.Name) ? snapshot.ChartType.ToString() : snapshot.Name;
        }

        private static bool IsPdfSupportedWorksheetImage(ExcelImage image, out string reason) {
            if (!IsPdfSupportedImageContentType(image.ContentType)) {
                reason = $"content type '{image.ContentType}' is not supported by the first-party PDF image writer.";
                return false;
            }

            if (image.WidthPixels <= 0 || image.HeightPixels <= 0) {
                reason = "image has non-positive dimensions.";
                return false;
            }

            byte[] bytes = image.GetBytes();
            if (bytes.Length == 0) {
                reason = "image has empty bytes.";
                return false;
            }

            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo imageInfo)) {
                reason = "image bytes do not contain a supported image header.";
                return false;
            }

            if (IsPngContentType(image.ContentType) && imageInfo.Format != OfficeImageFormat.Png) {
                reason = $"image bytes were declared as PNG but detected as {imageInfo.Format}.";
                return false;
            }

            if (IsJpegContentType(image.ContentType) && imageInfo.Format != OfficeImageFormat.Jpeg) {
                reason = $"image bytes were declared as JPEG but detected as {imageInfo.Format}.";
                return false;
            }

            reason = string.Empty;
            return true;
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return IsPngContentType(contentType) || IsJpegContentType(contentType);
        }

        private static bool IsPngContentType(string contentType) {
            return string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsJpegContentType(string contentType) {
            return string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(contentType, "image/jpg", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHyperlinkRelationship(ExternalRelationship relationship) {
            return relationship.RelationshipType.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase);
        }
    }
}
