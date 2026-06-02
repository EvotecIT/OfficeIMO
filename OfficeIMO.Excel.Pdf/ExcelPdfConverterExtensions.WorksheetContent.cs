using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetPageSetup(PdfCore.PdfPageCompose page, ExcelSheetPageSetup? pageSetup, ExcelPdfSaveOptions options) {
            PdfCore.PageSize? pageSize = options.PageSize;
            if (!options.PageSize.HasValue && pageSetup?.Orientation == ExcelPageOrientation.Landscape) {
                pageSize = (pageSize ?? PdfCore.PageSizes.Letter).Landscape();
            } else if (!options.PageSize.HasValue && pageSetup?.Orientation == ExcelPageOrientation.Portrait) {
                pageSize = (pageSize ?? PdfCore.PageSizes.Letter).Portrait();
            }

            if (pageSize.HasValue) {
                page.Size(pageSize.Value);
            }

            if (options.Margins.HasValue) {
                page.Margin(options.Margins.Value);
            } else if (pageSetup?.Margins != null) {
                page.Margin(ToPdfMargins(pageSetup.Margins));
            }
        }

        private static PdfCore.PageMargins ToPdfMargins(ExcelSheetPageMargins margins) {
            return PdfCore.PageMargins.FromInches(margins.Left, margins.Top, margins.Right, margins.Bottom);
        }

        private static IReadOnlyList<string> GetSheetNames(ExcelDocumentReader reader, ExcelPdfSaveOptions options) {
            IReadOnlyList<string> requestedNames = options.SheetNames ?? Array.Empty<string>();
            if (requestedNames.Count == 0) {
                return reader.GetSheetNames();
            }

            var names = new List<string>(requestedNames.Count);
            foreach (string name in requestedNames) {
                if (string.IsNullOrWhiteSpace(name)) {
                    throw new ArgumentException("Sheet names cannot contain null, empty, or whitespace values.", nameof(options));
                }

                ExcelSheetReader sheet = reader.GetSheet(name);
                names.Add(sheet.Name);
            }

            return names;
        }

        private static bool HasExplicitSheetSelection(ExcelPdfSaveOptions options) {
            return options.SheetNames != null && options.SheetNames.Count > 0;
        }

        private static bool ShouldSkipWorkbookSheet(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, bool hasExplicitSheetSelection) {
            return !hasExplicitSheetSelection
                   && options.RespectWorkbookSheetVisibility
                   && workbookSheet?.Hidden == true;
        }

        private static IReadOnlyList<WorksheetImageExportData> ReadWorksheetImages(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, string sheetName) {
            if (!options.UseWorksheetImages || workbookSheet == null) {
                return Array.Empty<WorksheetImageExportData>();
            }

            var images = new List<WorksheetImageExportData>();
            foreach (ExcelImage image in workbookSheet.Images.OrderBy(image => image.RowIndex).ThenBy(image => image.ColumnIndex)) {
                if (!IsPdfSupportedImageContentType(image.ContentType)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because content type '{image.ContentType}' is not supported by the first-party PDF image writer.");
                    continue;
                }

                byte[] bytes = image.GetBytes();
                if (bytes.Length == 0 || image.WidthPixels <= 0 || image.HeightPixels <= 0) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because it has empty image bytes or non-positive dimensions.");
                    continue;
                }

                if (!TryValidatePdfImageBytes(bytes, image.ContentType, out string? unsupportedReason)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because the first-party PDF image writer cannot export the image bytes. {unsupportedReason}");
                    continue;
                }

                images.Add(new WorksheetImageExportData(bytes, PixelsToPoints(image.WidthPixels), PixelsToPoints(image.HeightPixels), A1.CellReference(image.RowIndex, image.ColumnIndex)));
            }

            return images;
        }

        private static IReadOnlyList<WorksheetChartExportData> ReadWorksheetCharts(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, string sheetName) {
            if (!options.UseWorksheetCharts || workbookSheet == null) {
                return Array.Empty<WorksheetChartExportData>();
            }

            var charts = new List<WorksheetChartExportData>();
            foreach (ExcelChart chart in workbookSheet.Charts) {
                if (!chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{chart.Name}' was not exported because its chart data could not be read into a first-party PDF snapshot.");
                    continue;
                }

                if (IsSupportedChartSnapshot(snapshot)) {
                    charts.Add(new WorksheetChartExportData(snapshot));
                } else {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because chart type '{snapshot.ChartType}' is not supported by the first-party PDF chart snapshot renderer yet.");
                }
            }

            return charts
                .OrderBy(chart => chart.Snapshot.RowIndex)
                .ThenBy(chart => chart.Snapshot.ColumnIndex)
                .ToList();
        }

        private static bool IsSupportedChartSnapshot(ExcelChartSnapshot snapshot) {
            return IsColumnChart(snapshot.ChartType)
                   || IsBarChart(snapshot.ChartType)
                   || IsLineChart(snapshot.ChartType)
                   || IsAreaChart(snapshot.ChartType)
                   || IsScatterChart(snapshot.ChartType)
                   || IsRadarChart(snapshot.ChartType)
                   || IsPieChart(snapshot.ChartType)
                   || IsDoughnutChart(snapshot.ChartType);
        }

        private static string GetChartDisplayName(ExcelChartSnapshot snapshot) {
            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            return string.IsNullOrWhiteSpace(snapshot.Name) ? snapshot.ChartType.ToString() : snapshot.Name;
        }

        private static void AddWarning(ExcelPdfSaveOptions options, string sheetName, string feature, string message) {
            options.Warnings.Add(new ExcelPdfExportWarning(sheetName, feature, message));
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(contentType, "image/jpg", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryValidatePdfImageBytes(byte[] bytes, string contentType, out string? unsupportedReason) {
            unsupportedReason = null;
            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo imageInfo)) {
                unsupportedReason = "Image bytes do not contain a supported image header.";
                return false;
            }

            if (IsPngContentType(contentType) && imageInfo.Format != OfficeImageFormat.Png) {
                unsupportedReason = $"Image bytes were declared as PNG but were detected as {imageInfo.Format}.";
                return false;
            }

            if (IsJpegContentType(contentType) && imageInfo.Format != OfficeImageFormat.Jpeg) {
                unsupportedReason = $"Image bytes were declared as JPEG but were detected as {imageInfo.Format}.";
                return false;
            }

            try {
                _ = new PdfCore.PdfTableCellImage(bytes, 1D, 1D);
                return true;
            } catch (ArgumentException ex) {
                unsupportedReason = ex.Message;
                return false;
            } catch (InvalidDataException ex) {
                unsupportedReason = ex.Message;
                return false;
            } catch (NotSupportedException ex) {
                unsupportedReason = ex.Message;
                return false;
            }
        }

        private static bool IsPngContentType(string contentType) =>
            string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase);

        private static bool IsJpegContentType(string contentType) =>
            string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase)
            || string.Equals(contentType, "image/jpg", StringComparison.OrdinalIgnoreCase);

        private static double PixelsToPoints(int pixels) {
            return pixels * 72D / 96D;
        }

    }
}
