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
        private bool TryRegisterDeferredDirectDataSetImport(
            DataSet dataSet,
            bool createTables,
            TableStyle tableStyle,
            bool includeHeaders,
            bool includeAutoFilter,
            bool autoFit,
            CancellationToken ct,
            out IReadOnlyList<ExcelDataSetImportResult> results) {
            results = Array.Empty<ExcelDataSetImportResult>();
            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

            try {
                var model = DirectDataSetWorkbookModel.Create(
                    dataSet,
                    createTables,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    ct,
                    snapshotTables: true,
                    dateSystem: DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(dataSet, model, MaterializeDeferredDataSetImport, isDeferred: true, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = null;
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                results = model.Results;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        private void ClearDirectDataSetSaveCandidate() {
            var candidate = _directDataSetSaveCandidate;
            if (candidate == null) {
                _directDataSetMetadataSourceSheet = null;
                return;
            }

            _directDataSetSaveCandidate = null;
            _directDataSetMetadataSourceSheet = null;
            candidate.Dispose();
        }

        private bool TryBeginDeferredDirectSaveCandidateRegistration(bool replacingPendingDirectCellValues = false) {
            var candidate = _directDataSetSaveCandidate;
            if (candidate != null) {
                if (candidate.IsDeferred && candidate.IsValid) {
                    MaterializeDeferredDataSetImport();
                    return false;
                } else {
                    ClearDirectDataSetSaveCandidate();
                }
            }

            if (_pendingDirectCellValueSheet != null && !replacingPendingDirectCellValues) {
                MaterializePendingDirectCellValueSheetIfNeeded();
                return false;
            }

            return true;
        }

        internal bool TryReservePendingDirectCellValueSheet(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            if (_pendingDirectCellValueSheet == null) {
                _pendingDirectCellValueSheet = sheet;
                return true;
            }

            return ReferenceEquals(_pendingDirectCellValueSheet, sheet);
        }

        internal void ClearPendingDirectCellValueSheet(ExcelSheet sheet) {
            if (ReferenceEquals(_pendingDirectCellValueSheet, sheet)) {
                _pendingDirectCellValueSheet = null;
            }
        }

        private ExcelSheet? MaterializePendingDirectCellValueSheetIfNeeded() {
            var sheet = _pendingDirectCellValueSheet;
            if (sheet == null) {
                return null;
            }

            _pendingDirectCellValueSheet = null;
            sheet.MaterializePendingDirectCellValues();
            return sheet;
        }

        private void PromotePendingDirectCellValueSheetIfPossible() {
            var sheet = _pendingDirectCellValueSheet;
            if (sheet == null) {
                return;
            }

            if (!sheet.TryPromotePendingDirectCellValuesToSaveCandidate()) {
                MaterializePendingDirectCellValueSheetIfNeeded();
            }
        }

        internal void MaterializeDeferredDataSetImport() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                return;
            }

            _directDataSetSaveCandidate = null;
            candidate.Dispose();

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(candidate.Model);
            } finally {
                _materializingDeferredDataSetImport = false;
            }
        }

        internal bool HasDeferredDirectDataSetImport
            => !_materializingDeferredDataSetImport
               && _directDataSetSaveCandidate?.IsDeferred == true;

        internal bool HasPendingDirectCellValues => _pendingDirectCellValueSheet != null;

        private void MaterializeDirectDataSetModel(DirectDataSetWorkbookModel model) {
            foreach (var sheetModel in model.Sheets) {
                ExcelSheet sheet = TryGetExistingSheet(sheetModel.SheetName)
                    ?? AddWorkSheet(sheetModel.SheetName, SheetNameValidationMode.Strict);
                if (sheetModel.Range.Length == 0) {
                    continue;
                }

                DirectWorksheetMetadata? preservedMetadata = null;
                if (TryCaptureDirectWorksheetMetadata(sheetModel, out DirectWorksheetMetadata? sheetMetadata, out _, allowDrawings: true, allowUnsupportedOverlayStyles: true)) {
                    preservedMetadata = MergeDirectWorksheetMetadata(sheetModel.Metadata, sheetMetadata, replaceOverlayCells: true);
                }

                ResetWorksheetForDirectDataSetMaterialization(sheet.WorksheetPart);
                using var noLock = sheet.BeginNoLock();
                if (sheetModel.HasTable) {
                    string tableRange = sheet.InsertTabularRowSourceAsTableForDeferredMaterialization(
                        sheetModel.Table,
                        includeHeaders: sheetModel.IncludeHeaders,
                        tableName: sheetModel.TableName,
                        style: sheetModel.TableStyle,
                        includeAutoFilter: sheetModel.IncludeAutoFilter);
                    if (tableRange.Length == 0) {
                        tableRange = sheet.InsertDataTableAsTable(
                            sheetModel.Table.ToDataTable(),
                            includeHeaders: sheetModel.IncludeHeaders,
                            tableName: sheetModel.TableName,
                            style: sheetModel.TableStyle,
                            includeAutoFilter: sheetModel.IncludeAutoFilter);
                    }

                    if (tableRange.Length > 0) {
                        sheet.SetTableStyle(
                            tableRange,
                            sheetModel.TableStyle,
                            sheetModel.ShowFirstColumn,
                            sheetModel.ShowLastColumn,
                            sheetModel.ShowRowStripes,
                            sheetModel.ShowColumnStripes);
                    }
                } else {
                    if (!sheet.TryInsertTabularRowSourceForDeferredMaterialization(
                        sheetModel.Table,
                        includeHeaders: sheetModel.IncludeHeaders)) {
                        sheet.InsertDataTable(
                            sheetModel.Table.ToDataTable(),
                            includeHeaders: sheetModel.IncludeHeaders);
                    }
                }

                if (sheetModel.ColumnWidths is { Length: > 0 } columnWidths) {
                    sheet.ApplyAutoFitColumnWidthsForDeferredMaterialization(columnWidths);
                } else if (sheetModel.AutoFitColumns && sheetModel.Table.ColumnCount > 0) {
                    sheet.AutoFitColumnsFor(Enumerable.Range(1, sheetModel.Table.ColumnCount));
                }

                ApplyDirectMaterializedColumnNumberFormats(sheet, sheetModel);
                ApplyCapturedDirectWorksheetMetadata(sheet.WorksheetPart.Worksheet!, preservedMetadata, DateSystem);
            }
        }

        private static void ApplyDirectMaterializedColumnNumberFormats(ExcelSheet sheet, DirectDataSetSheetModel sheetModel) {
            var formats = sheetModel.ColumnNumberFormats;
            if (formats == null || formats.Count == 0 || sheetModel.Table.RowCount == 0) {
                return;
            }

            int startRow = sheetModel.IncludeHeaders ? 2 : 1;
            int endRow = startRow + sheetModel.Table.RowCount - 1;
            for (int i = 0; i < formats.Count && i < sheetModel.Table.ColumnCount; i++) {
                string? numberFormat = formats[i];
                if (string.IsNullOrWhiteSpace(numberFormat)) {
                    continue;
                }

                string column = A1.ColumnIndexToLetters(i + 1);
                sheet.FormatRange(
                    column + startRow.ToString(CultureInfo.InvariantCulture) + ":" + column + endRow.ToString(CultureInfo.InvariantCulture),
                    numberFormat!);
            }
        }

        private static void ApplyCapturedDirectWorksheetMetadata(Worksheet worksheet, DirectWorksheetMetadata? metadata, ExcelDateSystem dateSystem) {
            if (metadata == null || metadata.IsEmpty) {
                return;
            }

            if (!string.IsNullOrEmpty(metadata.SheetPropertiesXml)) {
                InsertWorksheetMetadataElement(worksheet, new SheetProperties(metadata.SheetPropertiesXml!), typeof(SheetDimension), typeof(SheetViews), typeof(SheetFormatProperties), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.SheetViewsXml)) {
                InsertWorksheetMetadataElement(worksheet, new SheetViews(metadata.SheetViewsXml!), typeof(SheetFormatProperties), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.SheetFormatPropertiesXml)) {
                InsertWorksheetMetadataElement(worksheet, CreateElementWithAttributes<SheetFormatProperties>(metadata.SheetFormatPropertiesXml!), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.SheetProtectionXml)) {
                InsertWorksheetMetadataElement(worksheet, CreateElementWithAttributes<SheetProtection>(metadata.SheetProtectionXml!), typeof(AutoFilter), typeof(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting), typeof(DataValidations), typeof(TableParts));
            }

            if (!string.IsNullOrEmpty(metadata.AutoFilterXml)) {
                InsertWorksheetMetadataElement(worksheet, new AutoFilter(metadata.AutoFilterXml!), typeof(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting), typeof(DataValidations), typeof(TableParts));
            }

            foreach (var conditionalFormattingXml in metadata.ConditionalFormattingXml) {
                InsertWorksheetMetadataElement(worksheet, new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting(conditionalFormattingXml), false, typeof(DataValidations), typeof(TableParts));
            }

            if (!string.IsNullOrEmpty(metadata.DataValidationsXml)) {
                InsertWorksheetMetadataElement(worksheet, new DataValidations(metadata.DataValidationsXml!), typeof(TableParts));
            }

            foreach (var xml in metadata.PostDataValidationXml) {
                var element = CreatePostDataValidationElement(xml);
                if (element != null) {
                    InsertWorksheetMetadataElement(worksheet, element, typeof(TableParts));
                }
            }

            if (!string.IsNullOrEmpty(metadata.DrawingXml)) {
                InsertWorksheetMetadataElement(worksheet, CreateElementWithAttributes<DocumentFormat.OpenXml.Spreadsheet.Drawing>(metadata.DrawingXml!), typeof(TableParts));
            }

            ApplyCapturedDirectOverlayCells(worksheet, metadata.OverlayCells, dateSystem);
        }

        private static void ApplyCapturedDirectOverlayCells(Worksheet worksheet, IReadOnlyList<DirectOverlayCell> overlayCells, ExcelDateSystem dateSystem) {
            if (overlayCells.Count == 0) {
                return;
            }

            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
            foreach (var overlayCell in overlayCells.OrderBy(static cell => cell.Row).ThenBy(static cell => cell.Column)) {
                if (overlayCell.IsDeleted) {
                    continue;
                }

                Row row = GetOrCreateDirectOverlayRow(sheetData, overlayCell.Row);
                Cell cell = GetOrCreateDirectOverlayCell(row, overlayCell.Row, overlayCell.Column);
                cell.StyleIndex = overlayCell.StyleIndex.HasValue ? overlayCell.StyleIndex.Value : null;
                ApplyCapturedDirectOverlayCellValue(cell, overlayCell.Value, dateSystem);
            }
        }

        private static Row GetOrCreateDirectOverlayRow(SheetData sheetData, int rowIndex) {
            Row? insertAfter = null;
            foreach (Row row in sheetData.Elements<Row>()) {
                uint currentIndex = row.RowIndex?.Value ?? 0U;
                if (currentIndex == (uint)rowIndex) {
                    return row;
                }

                if (currentIndex > (uint)rowIndex) {
                    break;
                }

                insertAfter = row;
            }

            var created = new Row { RowIndex = (uint)rowIndex };
            if (insertAfter == null) {
                var first = sheetData.Elements<Row>().FirstOrDefault();
                if (first == null) {
                    sheetData.Append(created);
                } else {
                    sheetData.InsertBefore(created, first);
                }
            } else if (insertAfter.NextSibling<Row>() == null) {
                sheetData.Append(created);
            } else {
                sheetData.InsertAfter(created, insertAfter);
            }

            return created;
        }

        private static Cell GetOrCreateDirectOverlayCell(Row row, int rowIndex, int columnIndex) {
            string reference = A1.CellReference(rowIndex, columnIndex);
            Cell? insertAfter = null;
            foreach (Cell cell in row.Elements<Cell>()) {
                if (string.Equals(cell.CellReference?.Value, reference, StringComparison.Ordinal)) {
                    return cell;
                }

                if (cell.CellReference?.Value is string currentReference
                    && currentReference.Length > 0
                    && GetDirectOverlayColumnIndex(currentReference) > columnIndex) {
                    break;
                }

                insertAfter = cell;
            }

            var created = new Cell { CellReference = reference };
            if (insertAfter == null) {
                var first = row.Elements<Cell>().FirstOrDefault();
                if (first == null) {
                    row.Append(created);
                } else {
                    row.InsertBefore(created, first);
                }
            } else if (insertAfter.NextSibling<Cell>() == null) {
                row.Append(created);
            } else {
                row.InsertAfter(created, insertAfter);
            }

            return created;
        }

        private static void ApplyCapturedDirectOverlayCellValue(Cell cell, object? value, ExcelDateSystem dateSystem) {
            cell.CellFormula = null;
            cell.InlineString = null;

            switch (value) {
                case null:
                case DBNull _:
                    cell.CellValue = new CellValue(string.Empty);
                    cell.DataType = CellValues.String;
                    break;
                case DirectFormulaCellValue formula:
                    cell.CellFormula = !string.IsNullOrEmpty(formula.FormulaXml)
                        ? CreateCellFormulaFromXml(formula.FormulaXml!)
                        : new CellFormula(formula.Formula);
                    cell.CellValue = formula.CachedValue != null ? new CellValue(formula.CachedValue) : null;
                    cell.DataType = null;
                    break;
                case DirectTypedCellValue typed:
                    cell.CellValue = typed.Value != null ? new CellValue(typed.Value) : null;
                    cell.DataType = GetDirectTypedCellDataType(typed.DataType);
                    cell.InlineString = !string.IsNullOrEmpty(typed.InlineStringXml)
                        ? CreateInlineStringFromXml(typed.InlineStringXml!)
                        : null;
                    break;
                case bool boolean:
                    cell.CellValue = new CellValue(boolean ? "1" : "0");
                    cell.DataType = CellValues.Boolean;
                    break;
                case byte number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case sbyte number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case short number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case ushort number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case int number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case uint number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case long number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case ulong number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case float number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case double number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case decimal number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case DateTime dateTime:
                    ApplyCapturedDirectOverlayNumber(cell, ExcelDateSystemConverter.ToSerial(dateTime, dateSystem));
                    break;
                default:
                    cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    cell.DataType = CellValues.String;
                    break;
            }
        }

        private static void ApplyCapturedDirectOverlayNumber<T>(Cell cell, T value) where T : IFormattable {
            cell.CellValue = new CellValue(value.ToString(null, CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }

        private static int GetDirectOverlayColumnIndex(string cellReference) {
            int column = 0;
            for (int i = 0; i < cellReference.Length; i++) {
                char ch = cellReference[i];
                if (ch >= 'A' && ch <= 'Z') {
                    column = checked((column * 26) + ch - 'A' + 1);
                } else if (ch >= 'a' && ch <= 'z') {
                    column = checked((column * 26) + ch - 'a' + 1);
                } else {
                    break;
                }
            }

            return column;
        }

        internal void MaterializeDeferredDataSetImportPreservingFastSaveModel() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                return;
            }

            DirectDataSetWorkbookModel? fastSaveModel = null;
            if (TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel? packageModel, out _, allowDrawings: true)) {
                fastSaveModel = packageModel;
            }

            _directDataSetSaveCandidate = null;
            candidate.Dispose();

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(candidate.Model);
                if (fastSaveModel != null) {
                    _materializedDirectDataSetFastSaveModel = fastSaveModel;
                    _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = true;
                    _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
                }
            } finally {
                _materializingDeferredDataSetImport = false;
            }
        }

        internal void PreserveDeferredDataSetFastSaveModelAndClearCandidate() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                ClearDirectDataSetSaveCandidate();
                return;
            }

            if (!TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel packageModel, out _, allowDrawings: true)) {
                return;
            }

            _materializedDirectDataSetFastSaveModel = packageModel;
            _materializedDirectDataSetFastSaveModelHasMaterializedWorksheet = false;
            _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
            _directDataSetSaveCandidate = null;
            candidate.Dispose();
        }

        private static void InsertWorksheetMetadataElement(Worksheet worksheet, DocumentFormat.OpenXml.OpenXmlElement element, params Type[] beforeTypes) {
            InsertWorksheetMetadataElement(worksheet, element, removeExistingSameType: true, beforeTypes);
        }

        private static void InsertWorksheetMetadataElement(Worksheet worksheet, DocumentFormat.OpenXml.OpenXmlElement element, bool removeExistingSameType, params Type[] beforeTypes) {
            if (removeExistingSameType) {
                foreach (var existing in worksheet.ChildElements.Where(child => child.GetType() == element.GetType()).ToList()) {
                    worksheet.RemoveChild(existing);
                }
            }

            foreach (var child in worksheet.ChildElements) {
                for (int i = 0; i < beforeTypes.Length; i++) {
                    if (beforeTypes[i].IsInstanceOfType(child)) {
                        worksheet.InsertBefore(element, child);
                        return;
                    }
                }
            }

            worksheet.Append(element);
        }

        private static DocumentFormat.OpenXml.OpenXmlElement? CreatePostDataValidationElement(string xml) {
            return GetXmlRootLocalName(xml) switch {
                "printOptions" => CreateElementWithAttributes<PrintOptions>(xml),
                "pageMargins" => CreateElementWithAttributes<PageMargins>(xml),
                "pageSetup" => CreateElementWithAttributes<PageSetup>(xml),
                "headerFooter" => new HeaderFooter(xml),
                "rowBreaks" => new RowBreaks(xml),
                "colBreaks" => new ColumnBreaks(xml),
                "cellWatches" => new CellWatches(xml),
                "ignoredErrors" => new DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors(xml),
                _ => null
            };
        }

        private static T CreateElementWithAttributes<T>(string xml) where T : DocumentFormat.OpenXml.OpenXmlElement, new() {
            var element = new T();
            using var reader = System.Xml.XmlReader.Create(new StringReader(xml), new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });

            if (!reader.Read() || reader.NodeType != System.Xml.XmlNodeType.Element) {
                return element;
            }

            if (reader.HasAttributes) {
                while (reader.MoveToNextAttribute()) {
                    if (reader.Prefix == "xmlns" || string.Equals(reader.Name, "xmlns", StringComparison.Ordinal)) {
                        continue;
                    }

                    element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        reader.Prefix,
                        reader.LocalName,
                        reader.NamespaceURI,
                        reader.Value));
                }
            }

            return element;
        }

        private static CellFormula CreateCellFormulaFromXml(string xml) {
            var formula = new CellFormula();
            using var reader = System.Xml.XmlReader.Create(new StringReader(xml), new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });

            if (!reader.Read() || reader.NodeType != System.Xml.XmlNodeType.Element) {
                return formula;
            }

            if (reader.HasAttributes) {
                while (reader.MoveToNextAttribute()) {
                    if (reader.Prefix == "xmlns" || string.Equals(reader.Name, "xmlns", StringComparison.Ordinal)) {
                        continue;
                    }

                    formula.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        reader.Prefix,
                        reader.LocalName,
                        reader.NamespaceURI,
                        reader.Value));
                }

                reader.MoveToElement();
            }

            formula.Text = reader.IsEmptyElement ? string.Empty : reader.ReadElementContentAsString();
            return formula;
        }

        private static InlineString CreateInlineStringFromXml(string xml) {
            try {
                return new InlineString(xml);
            } catch (ArgumentException) {
                return new InlineString();
            }
        }

        private static CellValues GetDirectTypedCellDataType(string dataType) {
            return dataType switch {
                "b" => CellValues.Boolean,
                "d" => CellValues.Date,
                "e" => CellValues.Error,
                "inlineStr" => CellValues.InlineString,
                "n" => CellValues.Number,
                "s" => CellValues.SharedString,
                "str" => CellValues.String,
                _ => CellValues.String
            };
        }

        private static string GetXmlRootLocalName(string xml) {
            if (string.IsNullOrWhiteSpace(xml)) {
                return string.Empty;
            }

            int start = 0;
            while (start < xml.Length && char.IsWhiteSpace(xml[start])) start++;
            if (start >= xml.Length || xml[start] != '<') {
                return string.Empty;
            }

            start++;
            if (start < xml.Length && xml[start] == '/') start++;
            int end = start;
            while (end < xml.Length && !char.IsWhiteSpace(xml[end]) && xml[end] != '>' && xml[end] != '/') end++;
            if (end <= start) {
                return string.Empty;
            }

            string qualifiedName = xml.Substring(start, end - start);
            int separator = qualifiedName.IndexOf(':');
            return separator >= 0 ? qualifiedName.Substring(separator + 1) : qualifiedName;
        }
    }
}
