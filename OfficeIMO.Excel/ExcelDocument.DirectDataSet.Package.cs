using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private ExcelSheet? TryGetExistingSheet(string sheetName) {
            Sheet? sheetElement = null;
            var sheets = WorkbookRoot.Sheets;
            if (sheets != null) {
                foreach (var candidate in sheets.Elements<Sheet>()) {
                    if (string.Equals(candidate.Name?.Value, sheetName, StringComparison.Ordinal)) {
                        sheetElement = candidate;
                        break;
                    }
                }
            }

            if (sheetElement?.Id == null) {
                return null;
            }

            if (WorkbookPartRoot.GetPartById(sheetElement.Id!) is not WorksheetPart) {
                return null;
            }

            return new ExcelSheet(this, _spreadSheetDocument, sheetElement);
        }

        private void ResetWorksheetForDirectDataSetMaterialization(WorksheetPart worksheetPart) {
            foreach (var tablePart in worksheetPart.TableDefinitionParts.ToList()) {
                string? tableName = tablePart.Table?.Name?.Value;
                if (!string.IsNullOrWhiteSpace(tableName)) {
                    RemoveReservedTableName(tableName!);
                }

                worksheetPart.DeletePart(tablePart);
            }

            worksheetPart.Worksheet = new Worksheet(new SheetData());
        }

        private bool TryWriteDirectDataSetPackage(
            Stream destination,
            ExcelSaveOptions? options,
            bool updateDocumentState,
            CancellationToken ct,
            out string? skipReason) {
            skipReason = null;

            if (destination == null || !destination.CanWrite) {
                skipReason = "Destination stream must be writable.";
                return false;
            }

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (options?.SafePreflight == true || options?.SafeRepairDefinedNames == true) {
                skipReason = "Save preflight options require the standard package finalization path.";
                return false;
            }

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save policy requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (WorkbookRoot.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Any() == true) {
                skipReason = "Workbook defined names require the standard package finalization path.";
                return false;
            }

            PromotePendingDirectCellValueSheetIfPossible();

            if (_requiresSavePreflight) {
                skipReason = "Save preflight is pending.";
                return false;
            }

            if (HasWorkbookContentOutsideDirectDataSetImport(allowSheets: true)) {
                skipReason = "Workbook-level metadata requires the standard package finalization path.";
                return false;
            }

            DirectDataSetWorkbookModel packageModel;
            if (_materializedDirectDataSetFastSaveModel != null) {
                if (!CanWriteMaterializedDirectDataSetPackage(_materializedDirectDataSetFastSaveModel)) {
                    skipReason = "Materialized direct DataSet content requires the extended package writer.";
                    return false;
                }

                if (!TryRefreshMaterializedDirectDataSetFastSaveModel(out skipReason)) {
                    skipReason ??= "Materialized direct DataSet fast-save metadata could not be refreshed.";
                    return false;
                }

                packageModel = _materializedDirectDataSetFastSaveModel!;
            } else {
                var candidate = _directDataSetSaveCandidate;
                if (candidate == null || !candidate.IsValid) {
                    skipReason = "No valid direct DataSet save candidate is available.";
                    ClearDirectDataSetSaveCandidate();
                    return false;
                }

                if (!TryCreateDirectPackageModel(candidate.Model, out packageModel, out skipReason)) {
                    return false;
                }
            }

            if (ct.CanBeCanceled) {
                ct.ThrowIfCancellationRequested();
            }

            PrepareDestinationStreamForWrite(destination);
            DirectDataSetWorkbookWriter.Write(destination, packageModel, ct);
            try { destination.Flush(); } catch (NotSupportedException) { }
            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
            }

            if (updateDocumentState) {
                _packageDirty = false;
                _packagePropertiesDirty = false;
                _requiresSavePreflight = false;
                _unchangedPackageBytes = null;
                _packageContentTypesKnownNormalized = true;
                _simplePackageContentKnown = true;
            }

            return true;
        }

        private bool CanWriteMaterializedDirectDataSetPackage(DirectDataSetWorkbookModel model) {
            for (int i = 0; i < model.Sheets.Count; i++) {
                var sheetModel = model.Sheets[i];
                if (!sheetModel.IncludeHeaders) {
                    return false;
                }

                ExcelSheet? sheet = TryGetExistingSheet(sheetModel.SheetName);
                if (sheet?.WorksheetPart.DrawingsPart != null) {
                    return false;
                }
            }

            return true;
        }

        private bool TryCreateDirectPackageModel(DirectDataSetWorkbookModel sourceModel, out DirectDataSetWorkbookModel model, out string? skipReason, bool allowDrawings = false) {
            DirectWorksheetMetadata?[]? metadata = null;
            for (int i = 0; i < sourceModel.Sheets.Count; i++) {
                var sheetModel = sourceModel.Sheets[i];
                if (!TryCaptureDirectWorksheetMetadata(sheetModel, out DirectWorksheetMetadata? sheetMetadata, out skipReason, allowDrawings)) {
                    model = sourceModel;
                    return false;
                }

                sheetMetadata = MergeDirectWorksheetMetadata(sheetModel.Metadata, sheetMetadata, replaceOverlayCells: true);
                if (sheetMetadata?.IsEmpty == false) {
                    metadata ??= new DirectWorksheetMetadata?[sourceModel.Sheets.Count];
                    metadata[i] = sheetMetadata;
                }
            }

            model = metadata != null ? sourceModel.WithWorksheetMetadata(metadata) : sourceModel;
            skipReason = null;
            return true;
        }

        private bool TryRefreshMaterializedDirectDataSetFastSaveModel(out string? skipReason) {
            skipReason = null;
            var model = _materializedDirectDataSetFastSaveModel;
            if (model == null) {
                return true;
            }

            if (!TryCreateDirectPackageModel(model, out DirectDataSetWorkbookModel refreshedModel, out skipReason, allowDrawings: true)) {
                return false;
            }

            _materializedDirectDataSetFastSaveModel = refreshedModel;
            return true;
        }

        private static DirectWorksheetMetadata? MergeDirectWorksheetMetadata(DirectWorksheetMetadata? existing, DirectWorksheetMetadata? captured, bool replaceOverlayCells = false) {
            if (existing == null || existing.IsEmpty) {
                return NormalizeDirectWorksheetMetadata(captured);
            }

            if (captured == null || captured.IsEmpty) {
                if (replaceOverlayCells && existing.OverlayCells.Count > 0) {
                    return NormalizeDirectWorksheetMetadata(new DirectWorksheetMetadata(
                        existing.SheetPropertiesXml,
                        existing.SheetViewsXml,
                        existing.SheetFormatPropertiesXml,
                        existing.SheetProtectionXml,
                        existing.AutoFilterXml,
                        existing.ConditionalFormattingXml,
                        existing.DataValidationsXml,
                        existing.DrawingXml,
                        existing.PostDataValidationXml,
                        Array.Empty<DirectOverlayCell>()));
                }

                return NormalizeDirectWorksheetMetadata(existing);
            }

            IReadOnlyList<DirectOverlayCell> overlayCells = replaceOverlayCells
                ? captured.OverlayCells
                : CombineOverlayCells(existing.OverlayCells, captured.OverlayCells);
            return NormalizeDirectWorksheetMetadata(new DirectWorksheetMetadata(
                existing.SheetPropertiesXml ?? captured.SheetPropertiesXml,
                existing.SheetViewsXml ?? captured.SheetViewsXml,
                existing.SheetFormatPropertiesXml ?? captured.SheetFormatPropertiesXml,
                existing.SheetProtectionXml ?? captured.SheetProtectionXml,
                existing.AutoFilterXml ?? captured.AutoFilterXml,
                CombineMetadataXmlLists(existing.ConditionalFormattingXml, captured.ConditionalFormattingXml),
                existing.DataValidationsXml ?? captured.DataValidationsXml,
                existing.DrawingXml ?? captured.DrawingXml,
                CombineMetadataXmlLists(existing.PostDataValidationXml, captured.PostDataValidationXml),
                overlayCells));
        }

        private static DirectWorksheetMetadata? NormalizeDirectWorksheetMetadata(DirectWorksheetMetadata? metadata) {
            if (metadata == null) {
                return null;
            }

            IReadOnlyList<DirectOverlayCell> overlayCells = metadata.OverlayCells;
            if (overlayCells.Count > 0) {
                List<DirectOverlayCell>? retainedOverlayCells = null;
                for (int i = 0; i < overlayCells.Count; i++) {
                    if (overlayCells[i].IsDeleted) {
                        retainedOverlayCells ??= new List<DirectOverlayCell>(overlayCells.Count);
                        for (int previous = 0; previous < i; previous++) {
                            retainedOverlayCells.Add(overlayCells[previous]);
                        }

                        continue;
                    }

                    retainedOverlayCells?.Add(overlayCells[i]);
                }

                if (retainedOverlayCells != null) {
                    metadata = new DirectWorksheetMetadata(
                        metadata.SheetPropertiesXml,
                        metadata.SheetViewsXml,
                        metadata.SheetFormatPropertiesXml,
                        metadata.SheetProtectionXml,
                        metadata.AutoFilterXml,
                        metadata.ConditionalFormattingXml,
                        metadata.DataValidationsXml,
                        metadata.DrawingXml,
                        metadata.PostDataValidationXml,
                        retainedOverlayCells.Count == 0 ? Array.Empty<DirectOverlayCell>() : retainedOverlayCells.ToArray());
                }
            }

            return metadata.IsEmpty ? null : metadata;
        }

        private static IReadOnlyList<string> CombineMetadataXmlLists(IReadOnlyList<string> first, IReadOnlyList<string> second) {
            if (first.Count == 0) return second;
            if (second.Count == 0) return first;

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var combined = new List<string>(first.Count + second.Count);
            for (int i = 0; i < first.Count; i++) {
                if (seen.Add(first[i])) {
                    combined.Add(first[i]);
                }
            }

            for (int i = 0; i < second.Count; i++) {
                if (seen.Add(second[i])) {
                    combined.Add(second[i]);
                }
            }

            return combined;
        }

        private static IReadOnlyList<DirectOverlayCell> CombineOverlayCells(IReadOnlyList<DirectOverlayCell> first, IReadOnlyList<DirectOverlayCell> second) {
            if (first.Count == 0) return second;
            if (second.Count == 0) return first;

            var combined = new Dictionary<(int Row, int Column), DirectOverlayCell>();
            for (int i = 0; i < first.Count; i++) {
                combined[(first[i].Row, first[i].Column)] = first[i];
            }

            for (int i = 0; i < second.Count; i++) {
                combined[(second[i].Row, second[i].Column)] = second[i];
            }

            return combined.Values
                .Where(static cell => !cell.IsDeleted)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray();
        }

        private bool TryCaptureDirectWorksheetMetadata(
            DirectDataSetSheetModel sheetModel,
            out DirectWorksheetMetadata? metadata,
            out string? skipReason,
            bool allowDrawings = false,
            bool allowUnsupportedOverlayStyles = false,
            ExcelSheet? metadataSourceOverride = null) {
            metadata = null;
            skipReason = null;

            ExcelSheet? sheet = null;
            if (metadataSourceOverride != null
                && ReferenceEquals(metadataSourceOverride.Document, this)
                && string.Equals(metadataSourceOverride.Name, sheetModel.SheetName, StringComparison.Ordinal)) {
                sheet = metadataSourceOverride;
            }

            var metadataSourceSheet = _directDataSetMetadataSourceSheet;
            if (sheet == null
                && metadataSourceSheet != null
                && ReferenceEquals(metadataSourceSheet.Document, this)
                && string.Equals(metadataSourceSheet.Name, sheetModel.SheetName, StringComparison.Ordinal)) {
                sheet = metadataSourceSheet;
            }

            sheet ??= TryGetExistingSheet(sheetModel.SheetName);
            if (sheet == null) {
                return true;
            }

            var worksheetPart = sheet.DeferredMetadataWorksheetPart;
            if (worksheetPart.DrawingsPart != null && !allowDrawings) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            if (worksheetPart.HyperlinkRelationships.Any()) {
                skipReason = "Worksheet contains hyperlink relationships.";
                return false;
            }

            foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                if (!sheetModel.HasTable) {
                    skipReason = "Worksheet contains table metadata outside the direct table model.";
                    return false;
                }

                var tableAutoFilter = tableDefinitionPart.Table?.Elements<AutoFilter>().FirstOrDefault();
                if (tableAutoFilter != null && tableAutoFilter.HasChildren) {
                    skipReason = "Worksheet contains table AutoFilter criteria outside the direct table model.";
                    return false;
                }
            }

            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null) {
                return true;
            }

            string? sheetPropertiesXml = null;
            string? sheetViewsXml = null;
            string? sheetFormatPropertiesXml = null;
            string? sheetProtectionXml = null;
            string? autoFilterXml = null;
            string? dataValidationsXml = null;
            string? drawingXml = null;
            IReadOnlyList<DirectOverlayCell> overlayCells = Array.Empty<DirectOverlayCell>();
            List<string>? conditionalFormattingXml = null;
            List<string>? postDataValidationXml = null;
            foreach (var child in worksheet.ChildElements) {
                switch (child) {
                    case SheetProperties sheetProperties when sheetPropertiesXml == null:
                        sheetPropertiesXml = sheetProperties.OuterXml;
                        break;
                    case SheetDimension:
                        break;
                    case SheetData sheetData:
                        if (!TryCaptureDirectWorksheetOverlayCells(sheet, sheetModel, sheetData, _spreadSheetDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet, allowUnsupportedOverlayStyles, out overlayCells, out skipReason)) {
                            return false;
                        }
                        break;
                    case SheetViews sheetViews when sheetViewsXml == null:
                        sheetViewsXml = sheetViews.OuterXml;
                        break;
                    case SheetFormatProperties sheetFormatProperties when sheetFormatPropertiesXml == null:
                        sheetFormatPropertiesXml = sheetFormatProperties.OuterXml;
                        break;
                    case SheetProtection sheetProtection when sheetProtectionXml == null:
                        sheetProtectionXml = sheetProtection.OuterXml;
                        break;
                    case Columns when sheetModel.ColumnWidths is { Length: > 0 }:
                        break;
                    case Columns:
                        skipReason = "Worksheet contains column metadata outside the direct DataSet column width model.";
                        return false;
                    case AutoFilter autoFilter when autoFilterXml == null:
                        autoFilterXml = autoFilter.OuterXml;
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting conditionalFormatting:
                        conditionalFormattingXml ??= new List<string>();
                        conditionalFormattingXml.Add(conditionalFormatting.OuterXml);
                        break;
                    case DataValidations dataValidations when dataValidationsXml == null:
                        dataValidationsXml = dataValidations.OuterXml;
                        break;
                    case PrintOptions:
                    case PageMargins:
                    case PageSetup:
                    case HeaderFooter:
                    case RowBreaks:
                    case ColumnBreaks:
                    case CellWatches:
                    case DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors:
                        postDataValidationXml ??= new List<string>();
                        postDataValidationXml.Add(child.OuterXml);
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.Drawing drawing when allowDrawings && drawingXml == null:
                        drawingXml = drawing.OuterXml;
                        break;
                    case TableParts when sheetModel.HasTable:
                        break;
                    default:
                        skipReason = "Worksheet contains unsupported element '" + child.LocalName + "' for the direct DataSet package writer.";
                        return false;
                }
            }

            if (sheetPropertiesXml == null
                && sheetViewsXml == null
                && sheetFormatPropertiesXml == null
                && sheetProtectionXml == null
                && autoFilterXml == null
                && dataValidationsXml == null
                && drawingXml == null
                && overlayCells.Count == 0
                && (conditionalFormattingXml == null || conditionalFormattingXml.Count == 0)
                && (postDataValidationXml == null || postDataValidationXml.Count == 0)) {
                return true;
            }

            metadata = new DirectWorksheetMetadata(
                sheetPropertiesXml,
                sheetViewsXml,
                sheetFormatPropertiesXml,
                sheetProtectionXml,
                autoFilterXml,
                conditionalFormattingXml?.ToArray() ?? Array.Empty<string>(),
                dataValidationsXml,
                drawingXml,
                postDataValidationXml?.ToArray() ?? Array.Empty<string>(),
                overlayCells);
            return true;
        }

        private static bool TryCaptureDirectWorksheetOverlayCells(
            ExcelSheet sheet,
            DirectDataSetSheetModel sheetModel,
            SheetData sheetData,
            Stylesheet? stylesheet,
            bool allowUnsupportedOverlayStyles,
            out IReadOnlyList<DirectOverlayCell> overlayCells,
            out string? skipReason) {
            overlayCells = Array.Empty<DirectOverlayCell>();
            skipReason = null;
            int directLastRow = sheetModel.Table.RowCount + (sheetModel.IncludeHeaders ? 1 : 0);
            List<DirectOverlayCell>? cells = null;
            Dictionary<uint, DirectOverlayStyleResolution>? styleResolutionCache = null;
            int nextRowIndex = 1;
            foreach (var row in sheetData.Elements<Row>()) {
                int rowIndex = row.RowIndex?.Value is uint explicitRow ? checked((int)explicitRow) : nextRowIndex;
                nextRowIndex = checked(rowIndex + 1);
                int nextColumnIndex = 1;
                foreach (var cell in row.Elements<Cell>()) {
                    if (!TryGetCellCoordinates(cell, rowIndex, nextColumnIndex, out int cellRow, out int cellColumn)) {
                        continue;
                    }

                    nextColumnIndex = checked(cellColumn + 1);
                    if (cellColumn <= 0 || (cellRow <= directLastRow && cellColumn <= sheetModel.Table.ColumnCount)) {
                        continue;
                    }

                    object? value = ReadDirectOverlayCellValue(sheet, cell);
                    if (value == null || value == DBNull.Value) {
                        cells ??= new List<DirectOverlayCell>();
                        cells.Add(new DirectOverlayCell(cellRow, cellColumn, null, null, null, isDeleted: true));
                        continue;
                    }

                    if (!TryResolveDirectOverlayNumberFormat(stylesheet, cell, allowUnsupportedOverlayStyles, ref styleResolutionCache, out string? numberFormat)) {
                        skipReason = "Worksheet contains overlay cell style metadata outside the direct DataSet style model.";
                        return false;
                    }

                    cells ??= new List<DirectOverlayCell>();
                    cells.Add(new DirectOverlayCell(cellRow, cellColumn, value, cell.StyleIndex?.Value, numberFormat));
                }
            }

            overlayCells = cells ?? (IReadOnlyList<DirectOverlayCell>)Array.Empty<DirectOverlayCell>();
            return true;
        }

        private static bool TryResolveDirectOverlayNumberFormat(
            Stylesheet? stylesheet,
            Cell cell,
            bool allowUnsupportedOverlayStyles,
            ref Dictionary<uint, DirectOverlayStyleResolution>? styleResolutionCache,
            out string? numberFormat) {
            numberFormat = null;
            if (cell.StyleIndex?.Value is not uint styleIndex) {
                return true;
            }

            styleResolutionCache ??= new Dictionary<uint, DirectOverlayStyleResolution>();
            if (!styleResolutionCache.TryGetValue(styleIndex, out var resolution)) {
                bool supported = TryResolveDirectOverlayStyle(stylesheet, styleIndex, allowUnsupportedOverlayStyles, out string? resolvedNumberFormat);
                resolution = new DirectOverlayStyleResolution(supported, resolvedNumberFormat);
                styleResolutionCache.Add(styleIndex, resolution);
            }

            numberFormat = resolution.NumberFormat;
            return resolution.Supported;
        }

        private static bool TryResolveDirectOverlayStyle(Stylesheet? stylesheet, uint styleIndex, bool allowUnsupportedOverlayStyles, out string? numberFormat) {
            numberFormat = null;
            if (styleIndex == 0U) {
                return true;
            }

            if (stylesheet == null) {
                return false;
            }

            var cellFormat = stylesheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
            if (cellFormat == null) {
                return allowUnsupportedOverlayStyles;
            }

            if (HasUnsupportedDirectOverlayStyle(cellFormat)) {
                return allowUnsupportedOverlayStyles;
            }

            if (cellFormat.NumberFormatId?.Value is not uint numberFormatId || numberFormatId == 0U) {
                return true;
            }

            string? customFormat = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
                ?.FormatCode
                ?.Value;
            numberFormat = customFormat ?? ResolveBuiltInNumberFormatCode(numberFormatId);
            return numberFormat != null;
        }

        private static bool HasUnsupportedDirectOverlayStyle(CellFormat cellFormat) {
            if ((cellFormat.FontId?.Value ?? 0U) != 0U
                || (cellFormat.FillId?.Value ?? 0U) != 0U
                || (cellFormat.BorderId?.Value ?? 0U) != 0U
                || (cellFormat.ApplyFont?.Value ?? false)
                || (cellFormat.ApplyFill?.Value ?? false)
                || (cellFormat.ApplyBorder?.Value ?? false)
                || (cellFormat.ApplyAlignment?.Value ?? false)
                || (cellFormat.ApplyProtection?.Value ?? false)
                || (cellFormat.QuotePrefix?.Value ?? false)
                || (cellFormat.PivotButton?.Value ?? false)
                || cellFormat.Alignment != null
                || cellFormat.Protection != null) {
                return true;
            }

            return false;
        }

        private static string? ResolveBuiltInNumberFormatCode(uint numberFormatId) {
            return numberFormatId switch {
                1U => "0",
                2U => "0.00",
                3U => "#,##0",
                4U => "#,##0.00",
                9U => "0%",
                10U => "0.00%",
                11U => "0.00E+00",
                12U => "# ?/?",
                13U => "# ??/??",
                14U => "mm-dd-yy",
                15U => "d-mmm-yy",
                16U => "d-mmm",
                17U => "mmm-yy",
                18U => "h:mm AM/PM",
                19U => "h:mm:ss AM/PM",
                20U => "h:mm",
                21U => "h:mm:ss",
                22U => "m/d/yy h:mm",
                37U => "#,##0 ;(#,##0)",
                38U => "#,##0 ;[Red](#,##0)",
                39U => "#,##0.00;(#,##0.00)",
                40U => "#,##0.00;[Red](#,##0.00)",
                45U => "mm:ss",
                46U => "[h]:mm:ss",
                47U => "mmss.0",
                48U => "##0.0E+0",
                49U => "@",
                _ => null
            };
        }

        private static bool TryGetCellCoordinates(Cell cell, int fallbackRow, int fallbackColumn, out int row, out int column) {
            row = 0;
            column = 0;
            string? reference = cell.CellReference?.Value;
            if (!string.IsNullOrWhiteSpace(reference)) {
                try {
                    (row, column) = A1.ParseCellRef(reference!);
                    return row > 0 && column > 0;
                } catch {
                    return false;
                }
            }

            row = fallbackRow;
            column = fallbackColumn;
            return row > 0 && column > 0;
        }

        private static object? ReadDirectOverlayCellValue(ExcelSheet sheet, Cell cell) {
            if (cell.CellFormula != null) {
                return new DirectFormulaCellValue(cell.CellFormula.Text ?? string.Empty, cell.CellFormula.OuterXml, cell.CellValue?.Text);
            }

            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;
            if (dataType == CellValues.Boolean) {
                return string.Equals(text, "1", StringComparison.Ordinal)
                       || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
            }

            if (dataType == null || dataType == CellValues.Number) {
                if (!string.IsNullOrWhiteSpace(text)) {
                    return new DirectTypedCellValue(cell.DataType?.InnerText ?? "n", text);
                }

                return text;
            }

            if (dataType == CellValues.Error || dataType == CellValues.Date || dataType == CellValues.InlineString) {
                string dataTypeText = cell.DataType?.InnerText
                                      ?? (dataType == CellValues.Error
                                          ? "e"
                                          : dataType == CellValues.Date
                                              ? "d"
                                              : "inlineStr");
                return new DirectTypedCellValue(dataTypeText, text, cell.InlineString?.OuterXml);
            }

            return sheet.GetCellText(cell);
        }

        private bool TrySaveDirectDataSetPackageToFile(string targetPath, ExcelSaveOptions? options, CancellationToken ct, out string? skipReason) {
            skipReason = null;
            var temporaryPath = CreateTemporarySavePath(targetPath);
            byte[]? packageBytes = null;

            try {
                using (var fs = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    if (!TryWriteDirectDataSetPackage(fs, options, updateDocumentState: false, ct, out skipReason)) {
                        return false;
                    }
                }

                packageBytes = File.ReadAllBytes(temporaryPath);

                try { _spreadSheetDocument.Dispose(); } catch { }
                ReplaceTargetFile(temporaryPath, targetPath);
                temporaryPath = string.Empty;
                ClearDirectDataSetSaveCandidate();
                ReloadFromBytes(packageBytes, simplePackageContentKnown: true);

                FilePath = targetPath;
                LastSaveDiagnostics = ExcelSaveDiagnostics.DirectDataSetPackage();
                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                skipReason = "Direct DataSet package writer failed: " + ex.Message;
                if (packageBytes != null) {
                    try {
                        ClearDirectDataSetSaveCandidate();
                        ReloadFromBytes(packageBytes, simplePackageContentKnown: true);
                    } catch {
                    }
                }

                return false;
            } finally {
                DeleteFileIfExists(temporaryPath);
            }
        }
    }
}
