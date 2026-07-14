using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool TryWriteSimpleWorkbookPackage(Stream destination, ExcelSaveOptions? options, bool updateDocumentState, out string? skipReason, CancellationToken ct = default) {
            skipReason = null;

            if (destination == null || !destination.CanWrite || !destination.CanSeek) {
                skipReason = "Destination stream must be writable and seekable.";
                return false;
            }

            ct.ThrowIfCancellationRequested();

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (_unchangedPackageBytes != null) {
                skipReason = "An unchanged package payload is already available.";
                return false;
            }

            if (_packageContentTypesKnownNormalized && !_simplePackageContentKnown) {
                skipReason = "Workbook was loaded or previously finalized; standard save preserves package metadata and relationships.";
                return false;
            }

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save work is pending.";
                return false;
            }

            if (WorkbookPartRoot.WorksheetParts.Any(static part => !part.IsRootElementLoaded)) {
                skipReason = "An unmaterialized worksheet requires the extended package writer.";
                return false;
            }

            if (!FastWorkbookPackageModel.TryCreate(_spreadSheetDocument, out var model, out string? modelSkipReason)) {
                skipReason = modelSkipReason ?? "Workbook contains parts or worksheet features outside the simple package writer surface.";
                return false;
            }

            ct.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            FastWorkbookPackageWriter.Write(destination, model, ct);

            destination.Flush();
            destination.Seek(0, SeekOrigin.Begin);
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

        private bool TryWriteExtendedWorkbookPackage(Stream destination, ExcelSaveOptions? options, bool updateDocumentState, out string? skipReason, CancellationToken ct = default) {
            skipReason = null;

            if (destination == null || !destination.CanWrite) {
                skipReason = "Destination stream must be writable.";
                return false;
            }

            ct.ThrowIfCancellationRequested();

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (_unchangedPackageBytes != null) {
                skipReason = "An unchanged package payload is already available.";
                return false;
            }

            if (_packageContentTypesKnownNormalized && !_simplePackageContentKnown) {
                skipReason = "Workbook was loaded or previously finalized; standard save preserves package metadata and relationships.";
                return false;
            }

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save work is pending.";
                return false;
            }

            Stopwatch? stageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            if (!TryRefreshMaterializedDirectDataSetFastSaveModel(out string? directModelSkipReason)) {
                skipReason = directModelSkipReason ?? "Direct worksheet metadata could not be refreshed.";
                return false;
            }

            if (!ExtendedWorkbookPackageModel.TryCreate(_spreadSheetDocument, _materializedDirectDataSetFastSaveModel, out var model, out string? modelSkipReason)) {
                skipReason = modelSkipReason ?? "Workbook contains parts outside the extended package writer surface.";
                return false;
            }
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.CreateModel");

            if (ShouldMaterializeMixedDirectWorkbookGlobalParts(model)
                && _materializedDirectDataSetFastSaveModel != null) {
                if (!_materializedDirectDataSetFastSaveModelHasMaterializedWorksheet) {
                    _materializingDeferredDataSetImport = true;
                    try {
                        MaterializeDirectDataSetModel(_materializedDirectDataSetFastSaveModel);
                    } finally {
                        _materializingDeferredDataSetImport = false;
                    }
                }

                _materializedDirectDataSetFastSaveModel = null;
                _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
                if (!ExtendedWorkbookPackageModel.TryCreate(_spreadSheetDocument, directDataSetModel: null, out model, out modelSkipReason)) {
                    skipReason = modelSkipReason ?? "Workbook contains parts outside the extended package writer surface after materializing direct data.";
                    return false;
                }

                ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.MaterializeMixedWorkbookDirectData");
            }

            ct.ThrowIfCancellationRequested();
            bool destinationBacksOpenPackage = ReferenceEquals(destination, _packageStream);
            using MemoryStream? stagedPackage = !destination.CanSeek || destinationBacksOpenPackage
                ? new MemoryStream()
                : null;
            Stream writeTarget = stagedPackage ?? destination;

            PrepareDestinationStreamForWrite(writeTarget);
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.PrepareDestination");
            ExtendedWorkbookPackageWriter.Write(writeTarget, model, ct, Execution);
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.WritePackage");

            writeTarget.Flush();
            if (stagedPackage != null) {
                stagedPackage.Position = 0;
                if (destinationBacksOpenPackage) {
                    PrepareDestinationStreamForWrite(destination);
                }

                stagedPackage.CopyTo(destination);
                destination.Flush();
                if (destination.CanSeek) {
                    destination.Seek(0, SeekOrigin.Begin);
                }
            } else {
                destination.Seek(0, SeekOrigin.Begin);
            }

            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.FlushAndSeek");
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

        private void ReportExtendedPackageTiming(Stopwatch? stopwatch, string operation) {
            if (stopwatch == null) {
                return;
            }

            Execution.ReportTiming(operation, stopwatch.Elapsed);
            stopwatch.Restart();
        }

        private static bool ShouldMaterializeMixedDirectWorkbookGlobalParts(ExtendedWorkbookPackageModel model) {
            if (model.DirectDataSetModel == null || model.DirectWorksheetModels.Count == 0) {
                return false;
            }

            int worksheetPartCount = 0;
            bool hasWorkbookGlobalPart = false;
            foreach (var part in model.Parts) {
                if (part.Part is WorksheetPart) {
                    worksheetPartCount++;
                } else if (part.Part is WorkbookStylesPart || part.Part is SharedStringTablePart) {
                    hasWorkbookGlobalPart = true;
                }
            }

            return hasWorkbookGlobalPart
                   && model.DirectWorksheetModels.Count < worksheetPartCount
                   && !CanUseMixedDirectWorksheetEntries(model);
        }

        private static bool CanUseMixedDirectWorksheetEntries(ExtendedWorkbookPackageModel model) {
            if (model.DirectDataSetModel == null
                || model.DirectDataSetModel.Sheets.Count == 0
                || model.DirectWorksheetModels.Count == 0
                || model.DirectWorksheetModels.Count != model.DirectDataSetModel.Sheets.Count) {
                return false;
            }

            int worksheetPartCount = 0;
            bool hasNonDirectWorksheet = false;
            bool nonDirectWorksheetUsesSharedStrings = false;
            foreach (var part in model.Parts) {
                if (part.Part is not WorksheetPart worksheetPart) {
                    continue;
                }

                worksheetPartCount++;
                if (model.DirectWorksheetModels.ContainsKey(worksheetPart)) {
                    continue;
                }

                hasNonDirectWorksheet = true;
                var worksheet = worksheetPart.Worksheet;
                if (worksheet == null
                    || !CanWriteSimpleWorksheet(worksheetPart, worksheet, out _, allowDrawings: true, allowPivotTables: true)
                    || WorksheetDependsOnWorkbookGlobalParts(worksheetPart, worksheet)) {
                    return false;
                }

                nonDirectWorksheetUsesSharedStrings |= WorksheetUsesSharedStrings(worksheet);
            }

            return hasNonDirectWorksheet
                   && (!nonDirectWorksheetUsesSharedStrings || ModelHasSharedStringPart(model))
                   && model.DirectWorksheetModels.Count < worksheetPartCount;
        }

        private static bool ModelHasSharedStringPart(ExtendedWorkbookPackageModel model)
            => model.Parts.Any(static part => part.Part is SharedStringTablePart);

        private static bool WorksheetDependsOnWorkbookGlobalParts(WorksheetPart worksheetPart, Worksheet worksheet) {
            if (worksheetPart.TableDefinitionParts.Any()
                || worksheetPart.PivotTableParts.Any()
                || worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>() != null) {
                return true;
            }

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                foreach (var column in columns.Elements<Column>()) {
                    if (column.Style != null) {
                        return true;
                    }
                }
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return false;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                if (row.CustomFormat?.Value == true || row.StyleIndex != null) {
                    return true;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.StyleIndex != null) {
                        return true;
                    }

                }
            }

            return false;
        }

        private static bool WorksheetUsesSharedStrings(Worksheet worksheet) {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return false;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.DataType?.Value == CellValues.SharedString) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static readonly System.Text.UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);
    }
}
