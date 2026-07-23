using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static bool IsHiddenSheet(DocumentFormat.OpenXml.Spreadsheet.Sheet sheet) {
            return sheet.State?.Value == DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.Hidden
                   || sheet.State?.Value == DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.VeryHidden;
        }

        private static bool HasWorkbookRecalculationRequest(DocumentFormat.OpenXml.Spreadsheet.Workbook workbook) {
            var properties = workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CalculationProperties>();
            return properties?.ForceFullCalculation?.Value == true
                   || properties?.FullCalculationOnLoad?.Value == true;
        }

        private static IEnumerable<string> DescribeWorkbookRecalculationRequest(DocumentFormat.OpenXml.Spreadsheet.Workbook workbook) {
            var properties = workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CalculationProperties>();
            if (properties?.ForceFullCalculation?.Value == true) {
                yield return "Workbook calculation properties set forceFullCalc.";
            }

            if (properties?.FullCalculationOnLoad?.Value == true) {
                yield return "Workbook calculation properties set fullCalcOnLoad.";
            }
        }

        private static void AddUnsupportedHeaderFooterImages(ExcelSheet.HeaderFooterSnapshot headerFooter, string sheetName, ref int count, List<string> details) {
            AddUnsupportedHeaderFooterImage(headerFooter.HeaderLeftImage, sheetName, "header left", ref count, details);
            AddUnsupportedHeaderFooterImage(headerFooter.HeaderCenterImage, sheetName, "header center", ref count, details);
            AddUnsupportedHeaderFooterImage(headerFooter.HeaderRightImage, sheetName, "header right", ref count, details);
            AddUnsupportedHeaderFooterImage(headerFooter.FooterLeftImage, sheetName, "footer left", ref count, details);
            AddUnsupportedHeaderFooterImage(headerFooter.FooterCenterImage, sheetName, "footer center", ref count, details);
            AddUnsupportedHeaderFooterImage(headerFooter.FooterRightImage, sheetName, "footer right", ref count, details);
        }

        private static void AddUnsupportedHeaderFooterImage(ExcelSheet.HeaderFooterImageSnapshot? image, string sheetName, string location, ref int count, List<string> details) {
            if (image == null || IsPdfSupportedHeaderFooterImage(image, out string reason)) {
                return;
            }

            count++;
            details.Add($"{sheetName} {location}: {reason}");
        }

        private static void AddUnsupportedPrintArea(ExcelSheet sheet, string sheetName, ref int count, List<string> details) {
            string? printArea = sheet.GetPrintArea();
            if (string.IsNullOrWhiteSpace(printArea) || !ContainsMultiplePrintAreas(printArea!)) {
                return;
            }

            count++;
            details.Add($"{sheetName}: {printArea} uses multiple print areas; the first-party PDF path exports the worksheet used range instead.");
        }

        private static void AddUnsupportedPrintTitles(ExcelSheet sheet, string sheetName, ref int count, List<string> details) {
            ExcelPrintTitles titles = sheet.GetPrintTitles();
            if (!titles.HasColumns) {
                return;
            }

            count++;
            details.Add($"{sheetName}: print-title columns {A1.ColumnIndexToLetters(titles.FirstColumn!.Value)}:{A1.ColumnIndexToLetters(titles.LastColumn!.Value)} are configured, but first-party PDF export repeats print-title rows only.");
        }

        private static bool ContainsMultiplePrintAreas(string printArea) {
            bool inQuotedSheetName = false;
            for (int i = 0; i < printArea.Length; i++) {
                char current = printArea[i];
                if (current == '\'') {
                    if (inQuotedSheetName && i + 1 < printArea.Length && printArea[i + 1] == '\'') {
                        i++;
                        continue;
                    }

                    inQuotedSheetName = !inQuotedSheetName;
                    continue;
                }

                if (current == ',' && !inQuotedSheetName) {
                    return true;
                }
            }

            return false;
        }

        private static void AddUnrenderedDrawingShapes(WorksheetPart worksheetPart, string sheetName, ref int count, List<string> details) {
            IReadOnlyList<ExcelWorksheetDrawingObjectInfo> drawings = ExcelWorksheetDrawingObjectResolver.FindDrawingObjects(worksheetPart);
            if (drawings.Count == 0) {
                return;
            }

            count += drawings.Count;
            foreach (ExcelWorksheetDrawingObjectInfo drawing in drawings) {
                string source = drawing.CellReference == null ? sheetName : sheetName + "!" + drawing.CellReference;
                details.Add($"{source}: {drawing.Name} ({drawing.Kind})");
            }
        }

        private static void AddUnsupportedWorksheetHyperlinks(
            WorkbookPart workbookPart,
            IReadOnlyList<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets,
            WorksheetPart worksheetPart,
            string sheetName,
            ref int count,
            List<string> details) {
            var hyperlinks = worksheetPart.Worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                return;
            }

            foreach (var hyperlink in hyperlinks.Elements<DocumentFormat.OpenXml.Spreadsheet.Hyperlink>()) {
                if (hyperlink.Id != null) {
                    continue;
                }

                string? location = hyperlink.Location?.Value;
                if (string.IsNullOrWhiteSpace(location)) {
                    continue;
                }

                if (!TryResolveInternalHyperlinkTarget(sheetName, location!, out string targetSheetName, out string targetCellReference)) {
                    count++;
                    details.Add($"{sheetName}: {hyperlink.Reference?.Value ?? "hyperlink"} -> {location} is not a worksheet cell target supported by the first-party PDF hyperlink writer.");
                    continue;
                }

                if (!IsDefaultPdfExportedCell(workbookPart, sheets, targetSheetName, targetCellReference, out string reason)) {
                    count++;
                    details.Add($"{sheetName}: {hyperlink.Reference?.Value ?? "hyperlink"} -> {location} is skipped by the first-party PDF hyperlink writer because {reason}");
                }
            }
        }

        private static void AddUnsupportedDrawingHyperlinks(WorksheetPart worksheetPart, string sheetName, ref int count, List<string> details) {
            foreach (OpenXmlPart part in EnumeratePartAndChildren(worksheetPart, new HashSet<Uri>())) {
                if (part == worksheetPart) {
                    continue;
                }

                foreach (HyperlinkRelationship relationship in part.HyperlinkRelationships) {
                    count++;
                    details.Add($"{sheetName}: {part.Uri} {relationship.Id} -> {relationship.Uri}");
                }

                foreach (ExternalRelationship relationship in part.ExternalRelationships.Where(IsHyperlinkRelationship)) {
                    count++;
                    details.Add($"{sheetName}: {part.Uri} {relationship.Id} -> {relationship.Uri}");
                }
            }
        }

        private static bool TryResolveInternalHyperlinkTarget(string currentSheetName, string location, out string sheetName, out string cellReference) {
            sheetName = currentSheetName;
            cellReference = string.Empty;

            string targetReference = location;
            if (SheetNameLookup.TryParseSheetQualifiedReference(location, out string parsedSheetName, out string parsedReference, allowExternalWorkbookReferences: false)) {
                sheetName = parsedSheetName;
                targetReference = parsedReference;
            }

            return TryGetTopLeftCellReference(targetReference, out cellReference);
        }

        private static bool IsDefaultPdfExportedCell(
            WorkbookPart workbookPart,
            IReadOnlyList<DocumentFormat.OpenXml.Spreadsheet.Sheet> sheets,
            string sheetName,
            string cellReference,
            out string reason) {
            var sheet = SheetNameLookup.FindByRequestedName(sheets, sheetName);
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.Id?.Value)) {
                reason = $"target sheet '{sheetName}' was not found.";
                return false;
            }

            if (IsHiddenSheet(sheet)) {
                reason = $"target sheet '{sheetName}' is hidden and skipped by default PDF export.";
                return false;
            }

            if (workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart) {
                reason = $"target sheet '{sheetName}' is not a worksheet.";
                return false;
            }

            if (!TryGetTopLeftCellReference(cellReference, out string normalizedCellReference)
                || !A1.TryParseCellReferenceFast(normalizedCellReference, out int targetRow, out int targetColumn)) {
                reason = $"target cell '{cellReference}' is not a supported A1 reference.";
                return false;
            }

            if (!TryGetWorksheetUsedRange(worksheetPart, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                reason = $"target sheet '{sheetName}' has no exported cells in the default PDF range.";
                return false;
            }

            if (targetRow < firstRow || targetRow > lastRow || targetColumn < firstColumn || targetColumn > lastColumn) {
                reason = $"target cell '{sheetName}!{normalizedCellReference}' is outside the default PDF exported range {A1.CellReference(firstRow, firstColumn)}:{A1.CellReference(lastRow, lastColumn)}.";
                return false;
            }

            reason = string.Empty;
            return true;
        }

        private static bool TryGetWorksheetUsedRange(WorksheetPart worksheetPart, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            firstRow = int.MaxValue;
            firstColumn = int.MaxValue;
            lastRow = 0;
            lastColumn = 0;

            foreach (var cell in worksheetPart.Worksheet?.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>() ?? Enumerable.Empty<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                string? reference = cell.CellReference?.Value;
                if (!A1.TryParseCellReferenceFast(reference, out int row, out int column)) {
                    continue;
                }

                firstRow = Math.Min(firstRow, row);
                firstColumn = Math.Min(firstColumn, column);
                lastRow = Math.Max(lastRow, row);
                lastColumn = Math.Max(lastColumn, column);
            }

            return lastRow > 0 && lastColumn > 0;
        }

        private static bool TryGetTopLeftCellReference(string referenceToken, out string cellReference) {
            cellReference = string.Empty;
            if (string.IsNullOrWhiteSpace(referenceToken)) {
                return false;
            }

            string token = referenceToken.Trim().Replace("$", string.Empty);
            if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out _, out _)) {
                cellReference = A1.CellReference(firstRow, firstColumn);
                return true;
            }

            if (!A1.TryParseCellReferenceFast(token, out int row, out int column)) {
                return false;
            }

            cellReference = A1.CellReference(row, column);
            return true;
        }

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

        private static string GetSafeChartDisplayName(ExcelChart chart) {
            try {
                if (!string.IsNullOrWhiteSpace(chart.Title)) {
                    return chart.Title!;
                }
            } catch {
                // Dangling chart relationships can make title lookup fail; fall back to frame metadata.
            }

            string name = chart.Name;
            if (!string.IsNullOrWhiteSpace(name)) {
                return name;
            }

            try {
                return chart.ChartType.ToString();
            } catch {
                return "unnamed chart";
            }
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

            byte[] bytes = image.ToBytes();
            if (bytes.Length == 0) {
                reason = "image has empty bytes.";
                return false;
            }

            if (!OfficeImagePdfCompatibility.TryValidate(bytes, out OfficeImageInfo? imageInfo, out string? unsupportedReason)) {
                reason = unsupportedReason ?? "image bytes are not supported by the first-party PDF image writer.";
                return false;
            }

            if (image.WidthPixels <= 0 || image.HeightPixels <= 0) {
                reason = "image has non-positive dimensions.";
                return false;
            }

            if (OfficeImagePdfCompatibility.TryGetSupportedContentTypeFormat(image.ContentType, out OfficeImageFormat declaredFormat) &&
                imageInfo!.Format != declaredFormat) {
                reason = $"image bytes were declared as {GetPdfImageFormatDisplayName(declaredFormat)} but detected as {imageInfo.Format}.";
                return false;
            }

            reason = string.Empty;
            return true;
        }

        private static bool IsPdfSupportedHeaderFooterImage(ExcelSheet.HeaderFooterImageSnapshot image, out string reason) {
            if (!IsPdfSupportedImageContentType(image.ContentType)) {
                reason = $"content type '{image.ContentType}' is not supported by the first-party PDF image writer.";
                return false;
            }

            if (image.WidthPoints <= 0 || image.HeightPoints <= 0) {
                reason = "image has non-positive dimensions.";
                return false;
            }

            if (image.Bytes.Length == 0) {
                reason = "image has empty bytes.";
                return false;
            }

            if (!OfficeImagePdfCompatibility.TryValidate(image.Bytes, out OfficeImageInfo? imageInfo, out string? unsupportedReason)) {
                reason = unsupportedReason ?? "image bytes are not supported by the first-party PDF image writer.";
                return false;
            }

            if (OfficeImagePdfCompatibility.TryGetSupportedContentTypeFormat(image.ContentType, out OfficeImageFormat declaredFormat) &&
                imageInfo!.Format != declaredFormat) {
                reason = $"image bytes were declared as {GetPdfImageFormatDisplayName(declaredFormat)} but detected as {imageInfo.Format}.";
                return false;
            }

            reason = string.Empty;
            return true;
        }

        private static bool IsSupportedPdfExternalHyperlink(Uri uri) {
            return uri.IsAbsoluteUri;
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return OfficeImagePdfCompatibility.IsSupportedContentType(contentType);
        }

        private static string GetPdfImageFormatDisplayName(OfficeImageFormat format) =>
            format == OfficeImageFormat.Jpeg ? "JPEG" : format.ToString().ToUpperInvariant();

        private static bool IsHyperlinkRelationship(ExternalRelationship relationship) {
            return relationship.RelationshipType.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase);
        }
    }
}
