using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetPageSetup(PdfCore.PdfPageCompose page, ExcelSheetPageSetup? pageSetup, ExcelPdfSaveOptions options) {
            if (ShouldApplyPageSize(options, pageSetup)) {
                page.Size(GetEffectivePageSize(options, pageSetup));
            }

            if (options.Margins.HasValue) {
                page.Margin(options.Margins.Value);
            } else if (pageSetup?.Margins != null) {
                page.Margin(ToPdfMargins(pageSetup.Margins));
            }
        }

        private static bool ShouldApplyPageSize(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            if (options.PageSize.HasValue) {
                return true;
            }

            if (options.PdfOptions != null) {
                return false;
            }

            return options.UseWorksheetPageSetup && pageSetup != null;
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

                byte[] bytes = image.ToBytes();
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

                string? alternativeText = string.IsNullOrWhiteSpace(image.Description)
                    ? string.IsNullOrWhiteSpace(image.Title) ? null : image.Title
                    : image.Description;
                images.Add(new WorksheetImageExportData(
                    bytes,
                    PixelsToPoints(image.WidthPixels),
                    PixelsToPoints(image.HeightPixels),
                    A1.CellReference(image.RowIndex, image.ColumnIndex),
                    image.RowIndex,
                    image.ColumnIndex,
                    PixelsToPoints(image.OffsetXPixels),
                    PixelsToPoints(image.OffsetYPixels),
                    image.RotationDegrees,
                    image.FlipHorizontal,
                    image.FlipVertical,
                    alternativeText));
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

                if (!HasRenderableChartData(snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because it does not contain renderable chart categories and series.");
                } else if (HasMixedSeriesChartTypes(snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because mixed per-series chart types are not supported by the first-party PDF chart snapshot renderer yet.");
                } else if (IsSupportedChartSnapshot(snapshot)) {
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
            return TryMapChartKind(snapshot.ChartType, out _);
        }

        private static bool HasRenderableChartData(ExcelChartSnapshot snapshot) {
            return snapshot.Data.Categories.Count > 0 && snapshot.Data.Series.Count > 0;
        }

        private static string GetChartDisplayName(ExcelChartSnapshot snapshot) {
            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            return string.IsNullOrWhiteSpace(snapshot.Name) ? snapshot.ChartType.ToString() : snapshot.Name;
        }

        private static void AddWarning(ExcelPdfSaveOptions options, string sheetName, string feature, string message) {
            var warning = new ExcelPdfExportWarning(sheetName, feature, message);
            options.Warnings.Add(warning);
            options.Report.Add(warning.ToConversionWarning());
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return OfficeImagePdfCompatibility.IsSupportedContentType(contentType);
        }

        private static bool TryValidatePdfImageBytes(byte[] bytes, string contentType, out string? unsupportedReason) {
            unsupportedReason = null;
            if (!OfficeImagePdfCompatibility.TryValidateDeclaredContentType(bytes, contentType, out _, out unsupportedReason)) {
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

        private static double PixelsToPoints(int pixels) {
            return pixels * 72D / 96D;
        }

    }
}
