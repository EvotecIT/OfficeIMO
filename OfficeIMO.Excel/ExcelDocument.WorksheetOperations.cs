using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly object WorksheetOperationTableIdLock = new object();

        /// <summary>
        /// Copies a worksheet within this workbook.
        /// </summary>
        /// <param name="sourceSheetName">Name of the worksheet to copy.</param>
        /// <param name="newSheetName">Requested name for the copied worksheet.</param>
        /// <param name="validationMode">How to validate or sanitize <paramref name="newSheetName"/>.</param>
        /// <returns>The copied worksheet.</returns>
        public ExcelSheet CopyWorksheet(string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode = SheetNameValidationMode.Sanitize) {
            return CopyWorksheet(GetSheet(sourceSheetName), newSheetName, validationMode);
        }

        /// <summary>
        /// Copies a worksheet within this workbook.
        /// </summary>
        /// <param name="sourceSheet">Worksheet to copy.</param>
        /// <param name="newSheetName">Requested name for the copied worksheet.</param>
        /// <param name="validationMode">How to validate or sanitize <paramref name="newSheetName"/>.</param>
        /// <returns>The copied worksheet.</returns>
        public ExcelSheet CopyWorksheet(ExcelSheet sourceSheet, string newSheetName, SheetNameValidationMode validationMode = SheetNameValidationMode.Sanitize) {
            return CopyWorksheetWithinWorkbook(sourceSheet, newSheetName, validationMode).Sheet;
        }

        private WorksheetPackageCopyResult CopyWorksheetWithinWorkbook(ExcelSheet sourceSheet, string newSheetName, SheetNameValidationMode validationMode) {
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));
            if (!ReferenceEquals(sourceSheet.Document, this)) {
                throw new ArgumentException("Source worksheet must belong to this workbook. Use CopyWorksheetFrom to copy between workbooks.", nameof(sourceSheet));
            }

            return Locking.ExecuteWrite(EnsureLock(), () => {
                string validatedName = ValidateOrSanitizeSheetName(newSheetName, validationMode, currentSheetName: null);
                WorksheetPart sourcePart = sourceSheet.WorksheetPart;
                WorksheetPart copiedPart = WorkbookPartRoot.AddNewPart<WorksheetPart>();
                copiedPart.Worksheet = (Worksheet)sourcePart.Worksheet!.CloneNode(true);
                RemoveRelationshipBackedWorksheetFeatures(copiedPart.Worksheet);
                Dictionary<string, string> tableNameMap = CopyWorksheetTables(sourcePart, copiedPart, rewriteCopiedTableReferences: true);
                copiedPart.Worksheet.Save();

                Sheet sheet = AppendWorksheetElement(copiedPart, validatedName);
                MarkSheetCacheDirty();
                WorkbookRoot.Save();
                return new WorksheetPackageCopyResult(new ExcelSheet(this, _spreadSheetDocument, sheet), tableNameMap);
            });
        }

        /// <summary>
        /// Copies a worksheet from another workbook into this workbook.
        /// </summary>
        /// <param name="sourceDocument">Workbook containing the source worksheet.</param>
        /// <param name="sourceSheetName">Name of the worksheet to copy.</param>
        /// <param name="newSheetName">Requested name for the copied worksheet.</param>
        /// <param name="validationMode">How to validate or sanitize <paramref name="newSheetName"/>.</param>
        /// <returns>The copied worksheet.</returns>
        public ExcelSheet CopyWorksheetFrom(ExcelDocument sourceDocument, string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode = SheetNameValidationMode.Sanitize) {
            return CopyWorksheetFrom(sourceDocument, sourceSheetName, newSheetName, validationMode, options: null);
        }

        /// <summary>
        /// Copies a worksheet from another workbook into this workbook.
        /// </summary>
        /// <param name="sourceDocument">Workbook containing the source worksheet.</param>
        /// <param name="sourceSheetName">Name of the worksheet to copy.</param>
        /// <param name="newSheetName">Requested name for the copied worksheet.</param>
        /// <param name="validationMode">How to validate or sanitize <paramref name="newSheetName"/>.</param>
        /// <param name="options">Copy strategy options.</param>
        /// <returns>The copied worksheet.</returns>
        public ExcelSheet CopyWorksheetFrom(ExcelDocument sourceDocument, string sourceSheetName, string newSheetName, SheetNameValidationMode validationMode, ExcelWorksheetCopyOptions? options) {
            if (sourceDocument == null) throw new ArgumentNullException(nameof(sourceDocument));
            if (string.IsNullOrWhiteSpace(sourceSheetName)) throw new ArgumentNullException(nameof(sourceSheetName));
            options ??= new ExcelWorksheetCopyOptions();
            if (options.MaxDefinedNames <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDefinedNames));
            if (options.MaxDefinedNameCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDefinedNameCharacters));
            if (ReferenceEquals(sourceDocument, this) && options.CopyMode != ExcelWorksheetCopyMode.Values) {
                return CopyWorksheet(sourceSheetName, newSheetName, validationMode);
            }

            return options.CopyMode == ExcelWorksheetCopyMode.Values
                ? CopyWorksheetFromValues(sourceDocument, sourceSheetName, newSheetName, validationMode)
                : CopyWorksheetFromPackage(
                    sourceDocument,
                    sourceSheetName,
                    newSheetName,
                    validationMode,
                    options,
                    new DefinedNameCopyBudget(options.MaxDefinedNames, options.MaxDefinedNameCharacters));
        }

        /// <summary>
        /// Reorders a worksheet by name using a zero-based target index.
        /// </summary>
        public void ReorderWorksheet(string sheetName, int targetIndex) {
            ReorderWorksheet(GetSheet(sheetName), targetIndex);
        }

        /// <summary>
        /// Reorders a worksheet using a zero-based target index.
        /// </summary>
        public void ReorderWorksheet(ExcelSheet sheet, int targetIndex) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                throw new ArgumentException("Worksheet must belong to this workbook.", nameof(sheet));
            }

            Locking.ExecuteWrite(EnsureLock(), () => {
                Sheets sheets = WorkbookRoot.Sheets ?? throw new InvalidOperationException("Workbook sheets collection is missing.");
                List<Sheet> orderedSheets = sheets.Elements<Sheet>().ToList();
                if (targetIndex < 0 || targetIndex >= orderedSheets.Count) {
                    throw new ArgumentOutOfRangeException(nameof(targetIndex), $"Index {targetIndex.ToString(CultureInfo.InvariantCulture)} is out of range (0..{(orderedSheets.Count - 1).ToString(CultureInfo.InvariantCulture)}).");
                }

                Sheet? target = orderedSheets.FirstOrDefault(s => ReferenceEquals(s, sheet.SheetElement)
                    || string.Equals(s.Name?.Value, sheet.Name, StringComparison.Ordinal));
                if (target == null) {
                    throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
                }

                target.Remove();
                orderedSheets.Remove(target);
                if (targetIndex >= orderedSheets.Count) {
                    sheets.Append(target);
                } else {
                    sheets.InsertBefore(target, orderedSheets[targetIndex]);
                }

                MarkSheetCacheDirty();
                WorkbookRoot.Save();
            });
        }

        /// <summary>
        /// Merges source worksheet rows into a target worksheet by appending or writing to a requested target position.
        /// This is a workbook operation over worksheet values; callers own any data shaping or relational joins before writing.
        /// </summary>
        public ExcelWorksheetMergeResult MergeWorksheets(ExcelSheet targetSheet, ExcelSheet sourceSheet, ExcelWorksheetMergeOptions? options = null) {
            if (targetSheet == null) throw new ArgumentNullException(nameof(targetSheet));
            if (sourceSheet == null) throw new ArgumentNullException(nameof(sourceSheet));
            if (!ReferenceEquals(targetSheet.Document, this)) {
                throw new ArgumentException("Target worksheet must belong to this workbook.", nameof(targetSheet));
            }

            options ??= new ExcelWorksheetMergeOptions();
            if (options.BlankRowsBefore < 0) throw new ArgumentOutOfRangeException(nameof(options.BlankRowsBefore));

            string sourceRange = string.IsNullOrWhiteSpace(options.SourceRange) ? sourceSheet.GetUsedRangeA1() : options.SourceRange!;
            var sourceBounds = ParseRangeOrCell(sourceRange);
            sourceRange = ToRange(sourceBounds);
            object?[,] values;
            using (var sourceReader = sourceSheet.Document.CreateReader()) {
                values = sourceReader.GetSheet(sourceSheet.Name).ReadRange(sourceRange);
            }

            bool skipHeader = options.SourceHasHeader && !options.IncludeSourceHeader && values.GetLength(0) > 0;
            int sourceRowOffset = skipHeader ? 1 : 0;
            int rowsToCopy = Math.Max(0, values.GetLength(0) - sourceRowOffset);
            return Locking.ExecuteWrite(EnsureLock(), () => {
                using (Locking.EnterNoLockScope()) {
                    int[] columnMap = BuildMergeColumnMap(targetSheet, sourceBounds, values, options);
                    int columnsToCopy = columnMap.Length;

                    int targetStartRow = options.TargetStartRow ?? GetAppendStartRow(targetSheet, options.BlankRowsBefore);
                    int targetStartColumn = GetMergeTargetStartColumn(targetSheet, sourceBounds, options);
                    if (targetStartRow < 1) throw new ArgumentOutOfRangeException(nameof(options.TargetStartRow));
                    if (targetStartColumn < 1) throw new ArgumentOutOfRangeException(nameof(options.TargetStartColumn));

                    EnsureMergeTargetCanWrite(targetSheet, targetStartRow, targetStartColumn, rowsToCopy, values, sourceRowOffset, columnMap, options);

                    var cells = new List<(int Row, int Column, object Value)>();
                    for (int rowOffset = 0; rowOffset < rowsToCopy; rowOffset++) {
                        for (int columnOffset = 0; columnOffset < columnsToCopy; columnOffset++) {
                            object? value = values[rowOffset + sourceRowOffset, columnMap[columnOffset]];
                            if (value == null) continue;
                            cells.Add((targetStartRow + rowOffset, targetStartColumn + columnOffset, value));
                        }
                    }

                    if (cells.Count > 0) {
                        targetSheet.CellValues(cells);
                    }

                    string targetRange = BuildMergeTargetRange(targetStartRow, targetStartColumn, rowsToCopy, columnsToCopy);
                    return new ExcelWorksheetMergeResult(
                        sourceSheet.Name,
                        targetSheet.Name,
                        sourceRange,
                        targetRange,
                        rowsToCopy,
                        columnsToCopy,
                        skipHeader);
                }
            });
        }

        /// <summary>
        /// Compares the used ranges of two worksheets.
        /// </summary>
        public IReadOnlyList<ExcelRangeDifference> CompareWorksheets(ExcelSheet leftSheet, ExcelSheet rightSheet, ExcelRangeCompareOptions? options = null) {
            if (leftSheet == null) throw new ArgumentNullException(nameof(leftSheet));
            if (rightSheet == null) throw new ArgumentNullException(nameof(rightSheet));
            return CompareRanges(leftSheet, leftSheet.GetUsedRangeA1(), rightSheet, rightSheet.GetUsedRangeA1(), options);
        }

        /// <summary>
        /// Compares two worksheet ranges and returns cell-level differences.
        /// </summary>
        public IReadOnlyList<ExcelRangeDifference> CompareRanges(
            ExcelSheet leftSheet,
            string leftRange,
            ExcelSheet rightSheet,
            string rightRange,
            ExcelRangeCompareOptions? options = null) {
            if (leftSheet == null) throw new ArgumentNullException(nameof(leftSheet));
            if (rightSheet == null) throw new ArgumentNullException(nameof(rightSheet));
            if (string.IsNullOrWhiteSpace(leftRange)) throw new ArgumentNullException(nameof(leftRange));
            if (string.IsNullOrWhiteSpace(rightRange)) throw new ArgumentNullException(nameof(rightRange));

            options ??= new ExcelRangeCompareOptions();
            var leftBounds = ParseRangeOrCell(leftRange);
            var rightBounds = ParseRangeOrCell(rightRange);
            object?[,] leftValues;
            object?[,] rightValues;

            using (var leftReader = leftSheet.Document.CreateReader()) {
                leftValues = leftReader.GetSheet(leftSheet.Name).ReadRange(ToRange(leftBounds));
            }

            using (var rightReader = rightSheet.Document.CreateReader()) {
                rightValues = rightReader.GetSheet(rightSheet.Name).ReadRange(ToRange(rightBounds));
            }

            int rows = Math.Max(leftValues.GetLength(0), rightValues.GetLength(0));
            int columns = Math.Max(leftValues.GetLength(1), rightValues.GetLength(1));
            var differences = new List<ExcelRangeDifference>();
            for (int rowOffset = 0; rowOffset < rows; rowOffset++) {
                for (int columnOffset = 0; columnOffset < columns; columnOffset++) {
                    bool hasLeft = rowOffset < leftValues.GetLength(0) && columnOffset < leftValues.GetLength(1);
                    bool hasRight = rowOffset < rightValues.GetLength(0) && columnOffset < rightValues.GetLength(1);
                    object? leftValue = hasLeft ? leftValues[rowOffset, columnOffset] : null;
                    object? rightValue = hasRight ? rightValues[rowOffset, columnOffset] : null;
                    if (hasLeft && hasRight && ValuesEqual(leftValue, rightValue, options)) {
                        continue;
                    }

                    int leftRow = leftBounds.r1 + rowOffset;
                    int leftColumn = leftBounds.c1 + columnOffset;
                    int rightRow = rightBounds.r1 + rowOffset;
                    int rightColumn = rightBounds.c1 + columnOffset;
                    ExcelRangeDifferenceKind kind = hasLeft
                        ? hasRight ? ExcelRangeDifferenceKind.ValueMismatch : ExcelRangeDifferenceKind.MissingFromRight
                        : ExcelRangeDifferenceKind.MissingFromLeft;
                    differences.Add(new ExcelRangeDifference(
                        kind,
                        leftRow,
                        leftColumn,
                        rightRow,
                        rightColumn,
                        A1.CellReference(leftRow, leftColumn),
                        A1.CellReference(rightRow, rightColumn),
                        leftValue,
                        rightValue));
                }
            }

            return differences;
        }

        private Sheet AppendWorksheetElement(WorksheetPart worksheetPart, string sheetName) {
            var workbook = WorkbookRoot;
            Sheets sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
            uint sheetId = ReadSheetElements()
                .Select(sheet => sheet.SheetId?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() + 1U;
            id.Add(sheetId);
            var sheet = new Sheet {
                Id = WorkbookPartRoot.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);
            return sheet;
        }

        private static void RemoveRelationshipBackedWorksheetFeatures(Worksheet worksheet) {
            worksheet.RemoveAllChildren<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            worksheet.RemoveAllChildren<LegacyDrawing>();
            worksheet.RemoveAllChildren<LegacyDrawingHeaderFooter>();
            worksheet.RemoveAllChildren<TableParts>();
            worksheet.RemoveAllChildren<OleObjects>();
            worksheet.RemoveAllChildren<Controls>();
            worksheet.RemoveAllChildren<Picture>();

            PageSetup? pageSetup = worksheet.GetFirstChild<PageSetup>();
            if (pageSetup != null) {
                pageSetup.Id = null;
            }

            RemoveElementsByLocalName(worksheet, "pivotTableDefinition");
            RemoveElementsByLocalName(worksheet, "pivotTableDefinitions");
            RemoveElementsByLocalName(worksheet, "queryTableParts");
            RemoveElementsByLocalName(worksheet, "customProperties");
            RemoveWorksheetExtensionsContainingLocalNames(worksheet, "slicerList", "slicerRef", "timelineRefs", "timelineRef");

            foreach (Hyperlinks hyperlinks in worksheet.Elements<Hyperlinks>().ToList()) {
                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>().Where(h => h.Id != null).ToList()) {
                    hyperlink.Remove();
                }

                if (!hyperlinks.Elements<Hyperlink>().Any()) {
                    hyperlinks.Remove();
                }
            }
        }

        private static void RemoveWorksheetExtensionsContainingLocalNames(Worksheet worksheet, params string[] localNames) {
            WorksheetExtensionList? extensionList = worksheet.GetFirstChild<WorksheetExtensionList>();
            if (extensionList == null) {
                return;
            }

            var names = new HashSet<string>(localNames, StringComparer.Ordinal);
            foreach (OpenXmlElement extension in extensionList.Elements<OpenXmlElement>().ToList()) {
                bool remove = extension.Descendants<OpenXmlElement>()
                    .Any(element => names.Contains(element.LocalName));
                if (remove) {
                    extension.Remove();
                }
            }

            if (!extensionList.Elements<OpenXmlElement>().Any()) {
                extensionList.Remove();
            }
        }

        private static void RemoveElementsByLocalName(OpenXmlElement root, string localName) {
            foreach (OpenXmlElement element in root.Descendants<OpenXmlElement>()
                .Where(element => string.Equals(element.LocalName, localName, StringComparison.Ordinal))
                .ToList()) {
                element.Remove();
            }
        }

        private static void RemoveExtensionsContainingLocalNames(OpenXmlElement root, params string[] localNames) {
            var names = new HashSet<string>(localNames, StringComparer.Ordinal);
            foreach (OpenXmlElement extension in root.Descendants<OpenXmlElement>()
                .Where(element => string.Equals(element.LocalName, "ext", StringComparison.Ordinal)
                    && element.Descendants<OpenXmlElement>().Any(descendant => names.Contains(descendant.LocalName)))
                .ToList()) {
                extension.Remove();
            }

            foreach (OpenXmlElement extensionList in root.Descendants<OpenXmlElement>()
                .Where(element => string.Equals(element.LocalName, "extLst", StringComparison.Ordinal)
                    && !element.ChildElements.Any())
                .ToList()) {
                extensionList.Remove();
            }
        }

        private Dictionary<string, string> CopyWorksheetTables(
            WorksheetPart sourcePart,
            WorksheetPart copiedPart,
            bool rewriteCopiedTableReferences = false,
            bool preserveTableFormulas = true) {
            TableParts? copiedTableParts = null;
            var tableNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var copiedTables = new List<Table>();
            foreach (TableDefinitionPart sourceTablePart in sourcePart.TableDefinitionParts) {
                Table? sourceTable = sourceTablePart.Table;
                if (sourceTable == null) {
                    continue;
                }

                string relationshipId = MakeUnusedRelationshipId(copiedPart);
                TableDefinitionPart copiedTablePart = copiedPart.AddNewPart<TableDefinitionPart>(relationshipId);
                var copiedTable = (Table)sourceTable.CloneNode(true);
                copiedTable.Id = GetNextUniqueTableId();
                StripCopiedTableQueryBindings(copiedTable);
                if (!preserveTableFormulas) {
                    StripCopiedTableFormulas(copiedTable);
                }
                string? sourceTableName = sourceTable.Name?.Value ?? sourceTable.DisplayName?.Value;
                string? sourceDisplayName = sourceTable.DisplayName?.Value;
                string tableName = CreateUniqueCopiedTableName(sourceTableName);
                copiedTable.Name = tableName;
                copiedTable.DisplayName = tableName;

                copiedTablePart.Table = copiedTable;
                ReserveTableName(tableName);
                copiedTables.Add(copiedTable);
                if (!string.IsNullOrWhiteSpace(sourceTableName)) {
                    tableNameMap[sourceTableName!] = tableName;
                }

                if (!string.IsNullOrWhiteSpace(sourceDisplayName)) {
                    tableNameMap[sourceDisplayName!] = tableName;
                }

                copiedTableParts ??= EnsureTableParts(copiedPart.Worksheet!);
                copiedTableParts.Append(new TablePart { Id = copiedPart.GetIdOfPart(copiedTablePart) });
            }

            if (rewriteCopiedTableReferences && tableNameMap.Count > 0) {
                RewriteStructuredTableReferences(copiedPart.Worksheet!, tableNameMap);
            }

            foreach (Table copiedTable in copiedTables) {
                if (rewriteCopiedTableReferences && tableNameMap.Count > 0) {
                    RewriteStructuredTableReferences(copiedTable, tableNameMap);
                }

                copiedTable.Save();
            }

            if (copiedTableParts != null) {
                copiedTableParts.Count = (uint)copiedTableParts.Elements<TablePart>().Count();
            }

            return tableNameMap;
        }

        private static void StripCopiedTableQueryBindings(Table table) {
            table.ConnectionId = null;
            foreach (TableColumn column in table.Descendants<TableColumn>()) {
                column.QueryTableFieldId = null;
            }

            RemoveElementsByLocalName(table, "queryTable");
            RemoveElementsByLocalName(table, "queryTableField");
            RemoveElementsByLocalName(table, "queryTableFields");
            RemoveExtensionsContainingLocalNames(table, "queryTable", "queryTableField", "queryTableFields");
        }

        private static void StripCopiedTableFormulas(Table table) {
            RemoveExtensionsContainingLocalNames(table, "calculatedColumnFormula", "totalsRowFormula");
            RemoveElementsByLocalName(table, "calculatedColumnFormula");
            RemoveElementsByLocalName(table, "totalsRowFormula");
        }

        private uint GetNextUniqueTableId() {
            lock (WorksheetOperationTableIdLock) {
                uint maxExistingId = 0;
                foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                    foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                        uint? id = tablePart.Table?.Id?.Value;
                        if (id.HasValue && id.Value > maxExistingId) {
                            maxExistingId = id.Value;
                        }
                    }
                }

                return maxExistingId + 1;
            }
        }

        private string CreateUniqueCopiedTableName(string? requestedName) {
            const int maxLength = 255;
            string baseName = SanitizeCopiedTableName(requestedName, maxLength);
            HashSet<string> used = GetOrInitTableNameCache();
            if (!used.Contains(baseName)) {
                return baseName;
            }

            int suffix = 2;
            while (true) {
                string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                int maxBaseLength = Math.Max(1, maxLength - suffixText.Length);
                string trimmedBase = baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName;
                string candidate = trimmedBase + suffixText;
                if (!used.Contains(candidate)) {
                    return candidate;
                }

                suffix++;
            }
        }

        private static string SanitizeCopiedTableName(string? requestedName, int maxLength) {
            string source = string.IsNullOrWhiteSpace(requestedName) ? "Table" : requestedName!;
            var sanitized = new System.Text.StringBuilder(source.Length);
            foreach (char ch in source) {
                sanitized.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
            }

            if (sanitized.Length == 0) {
                sanitized.Append("Table");
            }

            if (char.IsDigit(sanitized[0])) {
                sanitized.Insert(0, '_');
            }

            if (sanitized.Length > maxLength) {
                sanitized.Length = maxLength;
            }

            return sanitized.ToString();
        }

        private static TableParts EnsureTableParts(Worksheet worksheet) {
            TableParts? tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            if (tableParts != null) {
                return tableParts;
            }

            tableParts = new TableParts();
            WorksheetExtensionList? extensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
            if (extensionList != null) {
                worksheet.InsertBefore(tableParts, extensionList);
            } else {
                worksheet.Append(tableParts);
            }

            return tableParts;
        }

        private static void RewriteStructuredTableReferences(Worksheet worksheet, IReadOnlyDictionary<string, string> tableNameMap) {
            foreach (CellFormula formula in worksheet.Descendants<CellFormula>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }

            foreach (Formula formula in worksheet.Descendants<Formula>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }

            foreach (Formula1 formula in worksheet.Descendants<Formula1>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }

            foreach (Formula2 formula in worksheet.Descendants<Formula2>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }

            foreach (OfficeFormula formula in worksheet.Descendants<OfficeFormula>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }
        }

        private static void RewriteStructuredTableReferences(Table table, IReadOnlyDictionary<string, string> tableNameMap) {
            foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }

            foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>()) {
                formula.Text = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            }
        }

        private static void RewriteCopiedWorksheetExternalReferences(WorksheetPart worksheetPart, IReadOnlyDictionary<int, int> externalReferenceMap) {
            if (externalReferenceMap.Count == 0) {
                return;
            }

            foreach (CellFormula formula in worksheetPart.Worksheet!.Descendants<CellFormula>()) {
                formula.Text = RewriteExternalWorkbookReferenceIndexes(formula.Text, externalReferenceMap);
            }

            foreach (Formula formula in worksheetPart.Worksheet!.Descendants<Formula>()) {
                formula.Text = RewriteExternalWorkbookReferenceIndexes(formula.Text, externalReferenceMap);
            }

            foreach (Formula1 formula in worksheetPart.Worksheet!.Descendants<Formula1>()) {
                formula.Text = RewriteExternalWorkbookReferenceIndexes(formula.Text, externalReferenceMap);
            }

            foreach (Formula2 formula in worksheetPart.Worksheet!.Descendants<Formula2>()) {
                formula.Text = RewriteExternalWorkbookReferenceIndexes(formula.Text, externalReferenceMap);
            }

            foreach (OfficeFormula formula in worksheetPart.Worksheet!.Descendants<OfficeFormula>()) {
                formula.Text = RewriteExternalWorkbookReferenceIndexes(formula.Text, externalReferenceMap);
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                if (tablePart.Table == null) {
                    continue;
                }

                bool tableChanged = false;
                foreach (CalculatedColumnFormula formula in tablePart.Table.Descendants<CalculatedColumnFormula>()) {
                    tableChanged |= RewriteExternalWorkbookReferenceIndexFormula(formula, externalReferenceMap);
                }

                foreach (TotalsRowFormula formula in tablePart.Table.Descendants<TotalsRowFormula>()) {
                    tableChanged |= RewriteExternalWorkbookReferenceIndexFormula(formula, externalReferenceMap);
                }

                if (tableChanged) {
                    tablePart.Table.Save();
                }
            }
        }

        private static bool RewriteExternalWorkbookReferenceIndexFormula(OpenXmlLeafTextElement formula, IReadOnlyDictionary<int, int> externalReferenceMap) {
            string text = formula.Text;
            string rewritten = RewriteExternalWorkbookReferenceIndexes(text, externalReferenceMap);
            if (string.Equals(rewritten, text, StringComparison.Ordinal)) {
                return false;
            }

            formula.Text = rewritten;
            return true;
        }

        private static string RewriteExternalWorkbookReferenceIndexes(string? formula, IReadOnlyDictionary<int, int> externalReferenceMap) {
            if (string.IsNullOrEmpty(formula) || externalReferenceMap.Count == 0) {
                return formula ?? string.Empty;
            }

            var rewritten = new System.Text.StringBuilder(formula!.Length);
            bool inStringLiteral = false;
            for (int index = 0; index < formula!.Length;) {
                char current = formula[index];
                if (current == '"') {
                    rewritten.Append(current);
                    if (inStringLiteral && index + 1 < formula.Length && formula[index + 1] == '"') {
                        rewritten.Append(formula[index + 1]);
                        index += 2;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    index++;
                    continue;
                }

                if (!inStringLiteral
                    && current == '['
                    && TryReadExternalWorkbookIndex(formula, index, out int sourceIndex, out int closeIndex)
                    && LooksLikeExternalWorkbookReference(formula, index, closeIndex)
                    && externalReferenceMap.TryGetValue(sourceIndex, out int targetIndex)) {
                    rewritten.Append('[')
                        .Append(targetIndex.ToString(CultureInfo.InvariantCulture))
                        .Append(']');
                    index = closeIndex + 1;
                    continue;
                }

                rewritten.Append(current);
                index++;
            }

            return rewritten.ToString();
        }

        private static bool LooksLikeExternalWorkbookReference(string formula, int bracketIndex, int closeIndex) {
            if (bracketIndex > 0 && formula[bracketIndex - 1] == '[') {
                return false;
            }

            int nextIndex = closeIndex + 1;
            return nextIndex < formula.Length
                && (formula[nextIndex] == '\'' || char.IsLetterOrDigit(formula[nextIndex]) || formula[nextIndex] == '_');
        }

        private static bool TryReadExternalWorkbookIndex(string formula, int bracketIndex, out int sourceIndex, out int closeIndex) {
            sourceIndex = 0;
            closeIndex = bracketIndex;
            int index = bracketIndex + 1;
            if (index >= formula.Length || !char.IsDigit(formula[index])) {
                return false;
            }

            while (index < formula.Length && char.IsDigit(formula[index])) {
                int digit = formula[index] - '0';
                if (sourceIndex > (int.MaxValue - digit) / 10) {
                    return false;
                }

                sourceIndex = (sourceIndex * 10) + digit;
                index++;
            }

            if (index >= formula.Length || formula[index] != ']') {
                return false;
            }

            closeIndex = index;
            return sourceIndex > 0;
        }

        private static string RewriteStructuredTableReferences(string? formula, IReadOnlyDictionary<string, string> tableNameMap) {
            if (string.IsNullOrEmpty(formula) || tableNameMap.Count == 0) {
                return formula ?? string.Empty;
            }

            var mappings = tableNameMap
                .Where(mapping => !string.IsNullOrEmpty(mapping.Key))
                .OrderByDescending(mapping => mapping.Key.Length)
                .ToArray();
            if (mappings.Length == 0) {
                return formula!;
            }

            var rewritten = new System.Text.StringBuilder(formula!.Length);
            bool inStringLiteral = false;
            for (int index = 0; index < formula!.Length;) {
                char current = formula[index];
                if (current == '"') {
                    rewritten.Append(current);
                    if (inStringLiteral && index + 1 < formula.Length && formula[index + 1] == '"') {
                        rewritten.Append(formula[index + 1]);
                        index += 2;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    index++;
                    continue;
                }

                if (!inStringLiteral && TryRewriteStructuredReferenceAt(formula, index, mappings, out string? replacement, out int consumed)) {
                    rewritten.Append(replacement);
                    index += consumed;
                    continue;
                }

                rewritten.Append(current);
                index++;
            }

            return rewritten.ToString();
        }

        private static bool TryRewriteStructuredReferenceAt(
            string formula,
            int index,
            KeyValuePair<string, string>[] mappings,
            out string? replacement,
            out int consumed) {
            replacement = null;
            consumed = 0;
            foreach (var mapping in mappings) {
                string tableName = mapping.Key;
                if (index > 0) {
                    char previous = formula[index - 1];
                    if (char.IsLetterOrDigit(previous) || previous == '_' || previous == '\\' || previous == '\'' || previous == '!' || previous == '[') {
                        continue;
                    }
                }

                int nextIndex = index + tableName.Length;
                bool hasStructuredSpecifier = nextIndex < formula.Length && formula[nextIndex] == '[';
                bool hasBareReferenceBoundary = nextIndex >= formula.Length || IsFormulaTokenBoundary(formula[nextIndex]);
                if (!hasStructuredSpecifier && nextIndex < formula.Length && formula[nextIndex] == '(') {
                    continue;
                }

                if (!hasStructuredSpecifier && !hasBareReferenceBoundary) {
                    continue;
                }

                if (string.Compare(formula, index, tableName, 0, tableName.Length, StringComparison.OrdinalIgnoreCase) != 0) {
                    continue;
                }

                replacement = mapping.Value;
                consumed = tableName.Length;
                return true;
            }

            return false;
        }

        private static bool IsFormulaTokenBoundary(char value) {
            return !(char.IsLetterOrDigit(value) || value == '_' || value == '\\' || value == '\'' || value == '!' || value == ':' || value == '.');
        }

        private static string MakeUnusedRelationshipId(WorksheetPart worksheetPart) {
            var existing = new HashSet<string>(
                worksheetPart.Parts.Select(part => part.RelationshipId ?? string.Empty),
                StringComparer.Ordinal);
            int index = 1;
            string relationshipId;
            do {
                relationshipId = "rId" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }
            while (existing.Contains(relationshipId));

            return relationshipId;
        }

        private static (int r1, int c1, int r2, int c2) ParseRangeOrCell(string a1) {
            if (A1.TryParseRange(a1, out int r1, out int c1, out int r2, out int c2)) {
                return (r1, c1, r2, c2);
            }

            (int row, int column) = A1.ParseCellRef(a1);
            if (row > 0 && column > 0) {
                return (row, column, row, column);
            }

            throw new ArgumentException($"Invalid A1 range '{a1}'.", nameof(a1));
        }

        private static string ToRange((int r1, int c1, int r2, int c2) range) {
            return A1.CellReference(range.r1, range.c1) + ":" + A1.CellReference(range.r2, range.c2);
        }

        private static string BuildMergeTargetRange(int startRow, int startColumn, int rowCount, int columnCount) {
            int effectiveRows = Math.Max(1, rowCount);
            int effectiveColumns = Math.Max(1, columnCount);
            return A1.CellReference(startRow, startColumn) + ":" +
                A1.CellReference(startRow + effectiveRows - 1, startColumn + effectiveColumns - 1);
        }

        private static void EnsureMergeTargetCanWrite(
            ExcelSheet targetSheet,
            int targetStartRow,
            int targetStartColumn,
            int rowsToCopy,
            object?[,] sourceValues,
            int sourceRowOffset,
            IReadOnlyList<int> columnMap,
            ExcelWorksheetMergeOptions options) {
            if (options.OverwriteExistingCells || rowsToCopy == 0 || columnMap.Count == 0) {
                return;
            }

            var cellsToWrite = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int rowOffset = 0; rowOffset < rowsToCopy; rowOffset++) {
                for (int columnOffset = 0; columnOffset < columnMap.Count; columnOffset++) {
                    object? value = sourceValues[rowOffset + sourceRowOffset, columnMap[columnOffset]];
                    if (value == null) {
                        continue;
                    }

                    cellsToWrite.Add(A1.CellReference(targetStartRow + rowOffset, targetStartColumn + columnOffset));
                }
            }

            if (cellsToWrite.Count == 0) {
                return;
            }

            int targetEndRow = targetStartRow + rowsToCopy - 1;
            Worksheet? worksheet = targetSheet.WorksheetPart.Worksheet;
            SheetData? sheetData = worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return;
            }

            foreach (Row row in sheetData.Elements<Row>()) {
                if (row.RowIndex == null) {
                    continue;
                }

                int rowIndex = (int)row.RowIndex.Value;
                if (rowIndex < targetStartRow) {
                    continue;
                }

                if (rowIndex > targetEndRow) {
                    break;
                }

                foreach (Cell cell in row.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference) || !cellsToWrite.Contains(reference!)) {
                        continue;
                    }

                    if (CellHasContent(cell)) {
                        throw new InvalidOperationException($"Cannot merge into worksheet '{targetSheet.Name}' because cell {reference} already contains data. Set OverwriteExistingCells to true to replace existing values.");
                    }
                }
            }
        }

        private static bool CellHasContent(Cell cell) {
            if (cell.CellFormula != null) {
                return true;
            }

            if (cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text)) {
                return true;
            }

            return cell.InlineString != null;
        }

        private int[] BuildMergeColumnMap(
            ExcelSheet targetSheet,
            (int r1, int c1, int r2, int c2) sourceBounds,
            object?[,] sourceValues,
            ExcelWorksheetMergeOptions options) {
            int sourceColumnCount = sourceValues.GetLength(1);
            if (!options.MatchColumnsByHeader) {
                return Enumerable.Range(0, sourceColumnCount).ToArray();
            }

            if (!options.SourceHasHeader) {
                throw new ArgumentException("SourceHasHeader must be true when MatchColumnsByHeader is enabled.", nameof(options));
            }

            if (sourceValues.GetLength(0) == 0) {
                return Enumerable.Range(0, sourceColumnCount).ToArray();
            }

            int targetStartColumn = GetMergeTargetStartColumn(targetSheet, sourceBounds, options);
            int targetHeaderRow = GetMergeTargetHeaderRow(targetSheet, options);
            string targetHeaderRange = A1.CellReference(targetHeaderRow, targetStartColumn) + ":" +
                A1.CellReference(targetHeaderRow, targetStartColumn + sourceColumnCount - 1);
            object?[,] targetHeaders;
            using (var targetReader = targetSheet.Document.CreateReader()) {
                targetHeaders = targetReader.GetSheet(targetSheet.Name).ReadRange(targetHeaderRange);
            }

            var sourceColumns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int column = 0; column < sourceColumnCount; column++) {
                string header = NormalizeHeaderText(sourceValues[0, column]);
                if (header.Length == 0) {
                    throw new ArgumentException($"Source header at column {column + 1} is empty.", nameof(options));
                }

                if (sourceColumns.ContainsKey(header)) {
                    throw new ArgumentException($"Source header '{header}' appears more than once.", nameof(options));
                }

                sourceColumns.Add(header, column);
            }

            var columnMap = new int[sourceColumnCount];
            var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = 0; column < sourceColumnCount; column++) {
                string targetHeader = NormalizeHeaderText(targetHeaders[0, column]);
                if (targetHeader.Length == 0) {
                    throw new ArgumentException($"Target header at {A1.CellReference(targetHeaderRow, targetStartColumn + column)} is empty.", nameof(options));
                }

                if (!sourceColumns.TryGetValue(targetHeader, out int sourceColumn)) {
                    throw new ArgumentException($"Source range is missing column '{targetHeader}'.", nameof(options));
                }

                columnMap[column] = sourceColumn;
                matched.Add(targetHeader);
            }

            foreach (string sourceHeader in sourceColumns.Keys) {
                if (!matched.Contains(sourceHeader)) {
                    throw new ArgumentException($"Source column '{sourceHeader}' does not exist in the target header row.", nameof(options));
                }
            }

            return columnMap;
        }

        private int GetMergeTargetStartColumn(ExcelSheet targetSheet, (int r1, int c1, int r2, int c2) sourceBounds, ExcelWorksheetMergeOptions options) {
            if (options.TargetStartColumn.HasValue) {
                return options.TargetStartColumn.Value;
            }

            if (options.MatchColumnsByHeader) {
                var targetBounds = ParseRangeOrCell(targetSheet.GetUsedRangeA1());
                return targetBounds.c1;
            }

            return sourceBounds.c1;
        }

        private int GetMergeTargetHeaderRow(ExcelSheet targetSheet, ExcelWorksheetMergeOptions options) {
            if (options.TargetHeaderRow.HasValue) {
                if (options.TargetHeaderRow.Value < 1) throw new ArgumentOutOfRangeException(nameof(options.TargetHeaderRow));
                return options.TargetHeaderRow.Value;
            }

            if (options.TargetStartRow.HasValue) {
                if (options.TargetStartRow.Value <= 1) {
                    throw new ArgumentException("TargetHeaderRow must be set when TargetStartRow is 1 and MatchColumnsByHeader is enabled.", nameof(options));
                }

                return options.TargetStartRow.Value - 1;
            }

            var targetBounds = ParseRangeOrCell(targetSheet.GetUsedRangeA1());
            return targetBounds.r1;
        }

        private static string NormalizeHeaderText(object? value) {
            return System.Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        }

        private static int GetAppendStartRow(ExcelSheet targetSheet, int blankRowsBefore) {
            string usedRange = targetSheet.GetUsedRangeA1();
            var (_, _, endRow, _) = A1.ParseRange(usedRange);
            if (IsWorksheetEffectivelyEmpty(targetSheet, usedRange)) {
                return 1 + blankRowsBefore;
            }

            return endRow + 1 + blankRowsBefore;
        }

        private static bool IsWorksheetEffectivelyEmpty(ExcelSheet sheet, string usedRange) {
            using var reader = sheet.Document.CreateReader();
            object?[,] values = reader.GetSheet(sheet.Name).ReadRange(usedRange);
            for (int row = 0; row < values.GetLength(0); row++) {
                for (int column = 0; column < values.GetLength(1); column++) {
                    if (values[row, column] != null) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool ValuesEqual(object? left, object? right, ExcelRangeCompareOptions options) {
            object? normalizedLeft = NormalizeComparedValue(left, options);
            object? normalizedRight = NormalizeComparedValue(right, options);

            if (normalizedLeft is string leftText && normalizedRight is string rightText) {
                return string.Equals(leftText, rightText, options.IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
            }

            return object.Equals(normalizedLeft, normalizedRight);
        }

        private static object? NormalizeComparedValue(object? value, ExcelRangeCompareOptions options) {
            if (value == null) {
                return options.TreatNullAndEmptyStringAsEqual ? string.Empty : null;
            }

            if (value is string text) {
                string normalized = options.TrimStrings ? text.Trim() : text;
                return options.TreatNullAndEmptyStringAsEqual && normalized.Length == 0 ? string.Empty : normalized;
            }

            return value;
        }
    }
}
