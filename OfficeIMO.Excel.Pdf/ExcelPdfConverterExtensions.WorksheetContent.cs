using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetPageSetup(PdfCore.PdfPageBuilder page, ExcelSheetPageSetup? pageSetup, ExcelPdfExportOptions options) {
            if (ShouldApplyPageSize(options, pageSetup)) {
                page.Size(GetEffectivePageSize(options, pageSetup));
            }

            if (options.Margins.HasValue) {
                page.Margin(options.Margins.Value);
            } else if (pageSetup?.Margins != null) {
                page.Margin(ToPdfMargins(pageSetup.Margins));
            }
        }

        private static bool ShouldApplyPageSize(ExcelPdfExportOptions options, ExcelSheetPageSetup? pageSetup) {
            if (options.PageSize.HasValue) {
                return true;
            }

            if (options.HasExplicitPdfPageSizeConfiguration) {
                return false;
            }

            return options.UseWorksheetPageSetup && pageSetup != null;
        }

        private static PdfCore.PdfPageMargins ToPdfMargins(ExcelSheetPageMargins margins) {
            return new OfficeIMO.Pdf.PdfPageMargins(OfficeIMO.OfficeLength.Inches(margins.Left),OfficeIMO.OfficeLength.Inches(margins.Top),OfficeIMO.OfficeLength.Inches(margins.Right),OfficeIMO.OfficeLength.Inches(margins.Bottom));
        }

        private static IReadOnlyList<string> GetSheetNames(ExcelDocumentReader reader, ExcelPdfExportOptions options) {
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

        private static bool HasExplicitSheetSelection(ExcelPdfExportOptions options) {
            return options.SheetNames != null && options.SheetNames.Count > 0;
        }

        private static bool ShouldSkipWorkbookSheet(ExcelSheet? workbookSheet, ExcelPdfExportOptions options, bool hasExplicitSheetSelection) {
            return !hasExplicitSheetSelection
                   && options.RespectWorkbookSheetVisibility
                   && workbookSheet?.Hidden == true;
        }

        private static IReadOnlyList<WorksheetImageExportData> ReadWorksheetImages(ExcelSheet? workbookSheet, ExcelPdfExportOptions options, string sheetName) {
            if (!options.UseWorksheetImages || workbookSheet == null) {
                return Array.Empty<WorksheetImageExportData>();
            }

            var images = new List<WorksheetImageExportData>();
            foreach (ExcelImage image in workbookSheet.Images.OrderBy(image => image.RowIndex).ThenBy(image => image.ColumnIndex)) {
                if (!IsPdfSupportedImageContentType(image.ContentType)) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-image",
                        $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because content type '{image.ContentType}' is not supported by the first-party PDF image writer.");
                    continue;
                }

                byte[] bytes = image.ToBytes();
                if (bytes.Length == 0 || image.WidthPixels <= 0 || image.HeightPixels <= 0) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-image",
                        $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because it has empty image bytes or non-positive dimensions.");
                    continue;
                }

                if (!TryPreparePdfImageBytes(bytes, image.ContentType, out byte[] preparedBytes, out string? unsupportedReason)) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-image",
                        $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because the first-party PDF image writer cannot export the image bytes. {unsupportedReason}");
                    continue;
                }

                string? alternativeText = string.IsNullOrWhiteSpace(image.Description)
                    ? string.IsNullOrWhiteSpace(image.Title) ? null : image.Title
                    : image.Description;
                images.Add(new WorksheetImageExportData(
                    preparedBytes,
                    PixelsToPoints(image.WidthPixels),
                    PixelsToPoints(image.HeightPixels),
                    ExcelA1.CellReference(image.RowIndex, image.ColumnIndex),
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

        private static IReadOnlyList<WorksheetChartExportData> ReadWorksheetCharts(ExcelSheet? workbookSheet, ExcelPdfExportOptions options, string sheetName) {
            if (!options.UseWorksheetCharts || workbookSheet == null) {
                return Array.Empty<WorksheetChartExportData>();
            }

            AddWorksheetSmartArtWarnings(workbookSheet, options, sheetName);
            var charts = new List<WorksheetChartExportData>();
            foreach (ExcelChart chart in workbookSheet.Charts) {
                if (!chart.TryGetSnapshot(out ExcelChartInfo snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-chart",
                        $"Worksheet chart '{chart.Name}' was not exported because its chart data could not be read into a first-party PDF snapshot.");
                    continue;
                }

                foreach (var diagnostic in snapshot.Diagnostics.Where(diagnostic =>
                    diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported ||
                    diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated)) {
                    AddWarning(options, sheetName, diagnostic.Code, diagnostic.Message, diagnostic.LossKind);
                }

                if (!HasRenderableChartData(snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-chart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because it does not contain renderable chart categories and series.");
                } else if (IsSupportedChartSnapshot(snapshot, out string? unsupportedReason, out List<(string Code, string Message)> approximationWarnings)) {
                    charts.Add(new WorksheetChartExportData(snapshot));
                    foreach (var approximationWarning in approximationWarnings) {
                        AddWarning(
                            options,
                            sheetName,
                            approximationWarning.Code,
                            $"Worksheet chart '{GetChartDisplayName(snapshot)}' was exported with an explicit approximation: {approximationWarning.Message}");
                    }
                } else {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-chart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because {unsupportedReason}");
                }
            }

            return charts
                .OrderBy(chart => chart.Snapshot.RowIndex)
                .ThenBy(chart => chart.Snapshot.ColumnIndex)
                .ToList();
        }

        // Worksheet SmartArt has no first-party PDF renderer yet. Report each diagram rather than
        // dropping it silently.
        private static void AddWorksheetSmartArtWarnings(ExcelSheet workbookSheet, ExcelPdfExportOptions options, string sheetName) {
            var drawing = workbookSheet.WorksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) return;
            foreach (var diagram in drawing.Descendants<DocumentFormat.OpenXml.Drawing.Diagrams.RelationshipIds>()) {
                string? name = diagram.Ancestors<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault()?
                    .NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                AddWarning(
                    options,
                    sheetName,
                    "worksheet-smart-art",
                    $"Worksheet SmartArt diagram '{(string.IsNullOrWhiteSpace(name) ? "SmartArt" : name)}' was not exported because worksheet SmartArt has no first-party PDF renderer.");
            }
        }

        private static bool IsSupportedChartSnapshot(
            ExcelChartInfo snapshot,
            out string? unsupportedReason,
            out List<(string Code, string Message)> approximationWarnings) {
            unsupportedReason = null;
            approximationWarnings = new List<(string Code, string Message)>();
            if (!TryMapChartKind(snapshot.ChartType, out _)) {
                unsupportedReason = $"chart type '{snapshot.ChartType}' is not supported by the first-party PDF chart renderer.";
                return false;
            }
            // 3-D, pie-of-pie, stock and surface families are drawn as flat projections; say what is left out.
            if (ExcelRangeImageRenderer.TryMapChartKind(snapshot.ChartType, out _, out string? projection) &&
                !string.IsNullOrWhiteSpace(projection)) {
                approximationWarnings.Add((ExcelRangeImageRenderer.IsFlatChartProjection(snapshot.ChartType) ? OfficeDiagnosticCodes.ChartFlattened : OfficeDiagnosticCodes.ChartApproximation, projection!));
            }

            bool hasMixedSeries = snapshot.Data.Series.Any(series =>
                series.ChartType.HasValue && series.ChartType.Value != snapshot.ChartType);
            if (hasMixedSeries &&
                !ExcelRangeImageRenderer.TryMapSeriesRenderKind(snapshot.ChartType, out _, out _)) {
                unsupportedReason =
                    $"combo-chart base type '{snapshot.ChartType}' is not supported by the shared Cartesian combo-chart renderer.";
                return false;
            }

            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                OfficeChartType effectiveType = series.ChartType ?? snapshot.ChartType;
                if (effectiveType == snapshot.ChartType) {
                    continue;
                }

                if (!ExcelRangeImageRenderer.TryMapCompatibleComboSeriesRenderKind(
                        snapshot.ChartType,
                        effectiveType,
                        out _,
                        out string? approximation,
                        out string? incompatibility)) {
                    unsupportedReason = $"combo-chart series '{series.Name}' uses chart type '{effectiveType}': {incompatibility}";
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(approximation)) {
                    approximationWarnings.Add((ExcelRangeImageRenderer.IsFlatChartProjection(effectiveType) ? OfficeDiagnosticCodes.ChartFlattened : OfficeDiagnosticCodes.ChartApproximation, $"series '{series.Name}' ({effectiveType}): {approximation}"));
                }
            }

            return true;
        }

        private static bool HasRenderableChartData(ExcelChartInfo snapshot) {
            return snapshot.Data.Categories.Count > 0 && snapshot.Data.Series.Count > 0;
        }

        private static string GetChartDisplayName(ExcelChartInfo snapshot) {
            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            return string.IsNullOrWhiteSpace(snapshot.Name) ? snapshot.ChartType.ToString() : snapshot.Name;
        }

        private static void AddWarning(ExcelPdfExportOptions options, string sheetName, string feature, string message, OfficeConversionLossKind lossKind = OfficeConversionLossKind.Approximation) {
            var warning = new ExcelPdfExportWarning(sheetName, feature, message, lossKind);
            options.Warnings.Add(warning);
            options.Report.Add(warning.ToConversionWarning());
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return OfficeImagePdfCompatibility.IsSupportedContentType(contentType);
        }

        private static bool TryPreparePdfImageBytes(
            byte[] bytes,
            string contentType,
            out byte[] preparedBytes,
            out string? unsupportedReason) {
            preparedBytes = Array.Empty<byte>();
            unsupportedReason = null;
            if (!OfficeImagePdfCompatibility.TryValidateDeclaredContentType(bytes, contentType, out _, out unsupportedReason)) {
                return false;
            }

            return PdfCore.PdfDocument.TryPrepareImageBytes(
                bytes,
                out preparedBytes,
                out _,
                out _,
                out unsupportedReason);
        }

        private static double PixelsToPoints(int pixels) {
            return pixels * 72D / 96D;
        }

    }
}
