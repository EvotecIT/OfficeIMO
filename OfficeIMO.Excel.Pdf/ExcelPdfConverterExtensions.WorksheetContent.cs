using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetPageSetup(PdfCore.PdfPageBuilder page, ExcelSheetPageSetup? pageSetup, ExcelPdfExportOptions options) {
            if (ShouldApplyPageSize(options, pageSetup)) {
                page.Size(GetEffectivePageSize(options, pageSetup));
            }

            // The sheet scene, the header band and the page share one set of margins, so the
            // body and the running lines line up even when the workbook stores no margins.
            page.Margin(GetEffectiveMargins(options, pageSetup));
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
            return new OfficeIMO.Pdf.PdfPageMargins(margins.Left,margins.Top,margins.Right,margins.Bottom);
        }

        private static IReadOnlyList<string> GetSheetNames(ExcelDocument document, ExcelPdfExportOptions options) =>
            OfficeSheetSelector.ResolveMany(document, options.Sheets).Select(sheet => sheet.Name).ToArray();

        private static bool HasExplicitSheetSelection(ExcelPdfExportOptions options) {
            return options.Sheets != null && options.Sheets.Count > 0;
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
                string contentType = image.ContentType;
                byte[] bytes = image.ToBytes();
                byte[]? vectorBytes = null;
                // An SVG picture is drawn as vectors; one the PDF writer cannot draw exports its raster fallback.
                if (OfficeImageInfo.FromMimeType(contentType) == OfficeImageFormat.Svg) {
                    string? svgReason = "the SVG part is empty.";
                    if (bytes.Length > 0 && PdfCore.PdfDocument.TryPrepareImageKeepingSvg(bytes, default, out byte[] svg, out _, out svgReason)) {
                        vectorBytes = svg;
                    } else if (image.TryGetRasterFallback(out byte[] fallback, out string fallbackType)) {
                        bytes = fallback;
                        contentType = fallbackType;
                        AddWarning(
                            options,
                            sheetName,
                            OfficeDiagnosticCodes.SvgImageRasterFallback,
                            $"Worksheet SVG image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} could not be drawn as vectors, so its stored raster fallback was exported. {svgReason}");
                    }
                }

                if (vectorBytes == null && !IsPdfSupportedImageContentType(contentType)) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-image",
                        $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because content type '{contentType}' is not supported by the first-party PDF image writer.");
                    continue;
                }

                if (bytes.Length == 0 || image.Width <= OfficeLength.Zero || image.Height <= OfficeLength.Zero) {
                    AddWarning(
                        options,
                        sheetName,
                        "worksheet-image",
                        $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because it has empty image bytes or non-positive dimensions.");
                    continue;
                }

                byte[] preparedBytes = vectorBytes ?? Array.Empty<byte>();
                string? unsupportedReason = null;
                if (vectorBytes == null && !TryPreparePdfImageBytes(bytes, contentType, out preparedBytes, out unsupportedReason)) {
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
                    image.Width.ToPoints(),
                    image.Height.ToPoints(),
                    ExcelA1.CellReference(image.RowIndex, image.ColumnIndex),
                    image.RowIndex,
                    image.ColumnIndex,
                    image.OffsetX.ToPoints(),
                    image.OffsetY.ToPoints(),
                    image.RotationDegrees,
                    image.FlipHorizontal,
                    image.FlipVertical,
                    alternativeText,
                    ReadWorksheetImageCrop(image, options, sheetName),
                    ReadWorksheetImageLink(image, options, sheetName)));
            }

            return images;
        }

        // Like the Word and PowerPoint converters, pictures carry absolute link targets; a relative one is reported and omitted.
        private static string? ReadWorksheetImageLink(ExcelImage image, ExcelPdfExportOptions options, string sheetName) {
            Uri? link = image.HyperlinkUri;
            if (link == null) return null;
            if (link.IsAbsoluteUri) return link.AbsoluteUri;
            AddWarning(
                options,
                sheetName,
                "worksheet-image",
                $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} has a relative hyperlink target, so its PDF link was omitted.",
                OfficeConversionLossKind.Omission);
            return null;
        }

        // DrawingML crop ratios are fractions of the source; outsets or a collapsed rectangle fit the whole image instead.
        private static PdfCore.PdfImageSourceCrop? ReadWorksheetImageCrop(ExcelImage image, ExcelPdfExportOptions options, string sheetName) {
            double left = image.CropLeftRatio, top = image.CropTopRatio, right = image.CropRightRatio, bottom = image.CropBottomRatio;
            if (left == 0D && top == 0D && right == 0D && bottom == 0D) return null;
            if (left >= 0D && top >= 0D && right >= 0D && bottom >= 0D && left + right < 1D && top + bottom < 1D) {
                return new PdfCore.PdfImageSourceCrop(left, top, right, bottom);
            }

            AddWarning(
                options,
                sheetName,
                "worksheet-image",
                $"Worksheet image anchored at {ExcelA1.CellReference(image.RowIndex, image.ColumnIndex)} uses crop outsets or a collapsed crop rectangle; the whole image is fitted to its frame.",
                OfficeConversionLossKind.Approximation);
            return null;
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
