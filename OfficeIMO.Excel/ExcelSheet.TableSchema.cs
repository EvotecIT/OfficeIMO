using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Resolves a table name, display name, or current range to its stable package name.</summary>
        internal string ResolveTableName(string tableOrRange) {
            if (string.IsNullOrWhiteSpace(tableOrRange)) throw new ArgumentNullException(nameof(tableOrRange));
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                Table table = FindTableByRangeNameOrDisplayName(tableOrRange)
                    ?? throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                return table.Name?.Value ?? table.DisplayName?.Value
                    ?? throw new InvalidOperationException("Table name is missing.");
            });
        }

        /// <summary>Renames a table and every parsed structured reference to it across the workbook.</summary>
        public string RenameTable(
            string tableOrRange,
            string newName,
            TableNameValidationMode validationMode = TableNameValidationMode.Strict) {
            if (string.IsNullOrWhiteSpace(tableOrRange)) throw new ArgumentNullException(nameof(tableOrRange));
            string? result = null;
            WriteLock(() => {
                Table table = FindTableByRangeNameOrDisplayName(tableOrRange)
                    ?? throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                string oldName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                string oldDisplayName = table.DisplayName?.Value ?? oldName;
                bool stableNameMatches = string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase);
                if (stableNameMatches
                    && string.Equals(oldDisplayName, newName, StringComparison.OrdinalIgnoreCase)) {
                    result = oldName;
                    return;
                }
                string resolved = stableNameMatches
                    ? oldName
                    : EnsureValidUniqueTableName(newName, validationMode);
                table.Name = resolved;
                table.DisplayName = resolved;
                table.Save();
                _excelDocument.RemoveReservedTableName(oldName);
                _excelDocument.ReserveTableName(resolved);
                _excelDocument.RewriteTableNameReferences(
                    new[] { oldName, oldDisplayName },
                    resolved);
                WorksheetRoot.Save();
                _excelDocument.CleanupCalculationArtifacts(
                    save: false,
                    ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
                result = resolved;
            });
            return result!;
        }

        /// <summary>
        /// Replaces the ordered table-column schema and optionally resizes the table. The table's top-left
        /// cell remains fixed; callers retain control of data outside the new table rectangle.
        /// </summary>
        public IReadOnlyList<ExcelTableColumnInfo> SetTableSchema(
            string tableOrRange,
            IReadOnlyList<string> columnNames,
            string? newRange = null) => SetTableSchemaCore(
                tableOrRange,
                columnNames,
                newRange,
                retainExistingNames: false);

        private IReadOnlyList<ExcelTableColumnInfo> SetTableSchemaCore(
            string tableOrRange,
            IReadOnlyList<string>? columnNames,
            string? newRange,
            bool retainExistingNames) {
            if (string.IsNullOrWhiteSpace(tableOrRange)) throw new ArgumentNullException(nameof(tableOrRange));
            if (!retainExistingNames && columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            string[] names = retainExistingNames
                ? Array.Empty<string>()
                : ValidateTableColumnNames(columnNames!);
            IReadOnlyList<ExcelTableColumnInfo>? result = null;
            WriteLock(() => {
                Table table = FindTableByRangeNameOrDisplayName(tableOrRange)
                    ?? throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                var currentBounds = A1.ParseRange(table.Reference?.Value
                    ?? throw new InvalidOperationException("Table reference is missing."));
                var targetBounds = string.IsNullOrWhiteSpace(newRange)
                    ? currentBounds
                    : A1.ParseRange(newRange!);
                if (targetBounds.r1 != currentBounds.r1 || targetBounds.c1 != currentBounds.c1) {
                    throw new InvalidOperationException("Table schema resize must preserve the table's top-left cell. Move the cell range explicitly before resizing the table.");
                }
                int targetWidth = targetBounds.c2 - targetBounds.c1 + 1;
                if (retainExistingNames) {
                    names = CreateResizedTableColumnNames(table, targetWidth);
                }
                if (targetWidth != names.Length) {
                    throw new ArgumentException("The table range width must match the number of column names.", nameof(columnNames));
                }
                int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
                int totalsRows = (int)(table.TotalsRowCount?.Value ?? 0U);
                if (table.TotalsRowShown?.Value == true) totalsRows = Math.Max(1, totalsRows);
                if (targetBounds.r2 - targetBounds.r1 + 1 < Math.Max(1, headerRows + totalsRows)) {
                    throw new InvalidOperationException("The resized table must retain its configured header and totals rows.");
                }
                bool rangeChanged = targetBounds.r1 != currentBounds.r1
                    || targetBounds.c1 != currentBounds.c1
                    || targetBounds.r2 != currentBounds.r2
                    || targetBounds.c2 != currentBounds.c2;
                if (rangeChanged) {
                    string targetText = A1.CellReference(targetBounds.r1, targetBounds.c1) + ":" + A1.CellReference(targetBounds.r2, targetBounds.c2);
                    EnsureNoIntersectingOwnedStructures(
                        ExcelReference.Parse(targetText),
                        "Table resize would overlap another table, merged cells, an array or data-table formula, or PivotTable output.",
                        excludedTable: table);
                }

                TableColumns columns = table.TableColumns ??= new TableColumns();
                List<TableColumn> existing = columns.Elements<TableColumn>().ToList();
                string[] removedNames = existing.Skip(names.Length)
                    .Select(column => column.Name?.Value ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .ToArray();
                uint nextId = existing.Count == 0 ? 1U : existing.Max(column => column.Id?.Value ?? 0U) + 1U;
                var renames = new List<(string OldName, string NewName)>();
                for (int index = 0; index < names.Length; index++) {
                    TableColumn column;
                    if (index < existing.Count) {
                        column = existing[index];
                        string oldName = column.Name?.Value ?? string.Empty;
                        if (!string.Equals(oldName, names[index], StringComparison.Ordinal)) {
                            renames.Add((oldName, names[index]));
                        }
                    } else {
                        column = columns.AppendChild(new TableColumn { Id = nextId++ });
                    }
                    column.Name = names[index];
                }
                for (int index = existing.Count - 1; index >= names.Length; index--) existing[index].Remove();
                columns.Count = (uint)names.Length;

                string normalizedRange = A1.CellReference(targetBounds.r1, targetBounds.c1) + ":" + A1.CellReference(targetBounds.r2, targetBounds.c2);
                if (rangeChanged) RelocateTableTotalsRows(
                    currentBounds,
                    targetBounds,
                    totalsRows);
                table.Reference = normalizedRange;
                if (rangeChanged) RemapTableResizeSortReferences(
                    table,
                    currentBounds,
                    targetBounds,
                    headerRows,
                    totalsRows);
                AutoFilter? filter = table.GetFirstChild<AutoFilter>();
                if (filter != null) {
                    int filterLastRow = Math.Max(targetBounds.r1, targetBounds.r2 - totalsRows);
                    filter.Reference = A1.CellReference(targetBounds.r1, targetBounds.c1) + ":" + A1.CellReference(filterLastRow, targetBounds.c2);
                    foreach (FilterColumn stale in filter.Elements<FilterColumn>()
                        .Where(item => (item.ColumnId?.Value ?? uint.MaxValue) >= (uint)names.Length).ToList()) stale.Remove();
                }

                if (headerRows > 0) {
                    for (int index = 0; index < names.Length; index++) {
                        CellValueCoreNoMaterialize(targetBounds.r1, targetBounds.c1 + index, names[index]);
                    }
                }
                SynchronizeQueryTableSchema(table, columns, names);

                string tableName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                var renameMap = renames
                    .Where(item => item.OldName.Length > 0)
                    .ToDictionary(item => item.OldName, item => item.NewName, StringComparer.OrdinalIgnoreCase);
                if (removedNames.Length > 0) _excelDocument.InvalidateTableColumnReferences(tableName, removedNames, table);
                if (renameMap.Count > 0) _excelDocument.RewriteTableColumnReferences(tableName, renameMap, table);
                table.Save();
                WorksheetRoot.Save();
                bool schemaChanged = rangeChanged
                    || existing.Count != names.Length
                    || renames.Count > 0;
                if (schemaChanged) {
                    _excelDocument.CleanupCalculationArtifacts(
                        save: false,
                        ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
                }
                result = new ReadOnlyCollection<ExcelTableColumnInfo>(names
                    .Select((name, index) => new ExcelTableColumnInfo(index + 1, name))
                    .ToArray());
            });
            return result!;
        }

        /// <summary>Resizes a table while retaining its current ordered column names.</summary>
        public IReadOnlyList<ExcelTableColumnInfo> ResizeTable(string tableOrRange, string newRange) =>
            SetTableSchemaCore(
                tableOrRange,
                columnNames: null,
                newRange: newRange,
                retainExistingNames: true);

        private static string[] CreateResizedTableColumnNames(Table table, int targetWidth) {
            string[] names = table.TableColumns?.Elements<TableColumn>()
                .Select(column => column.Name?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>();
            if (targetWidth > names.Length) {
                var used = new HashSet<string>(names.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
                Array.Resize(ref names, targetWidth);
                for (int index = 0; index < names.Length; index++) {
                    if (string.IsNullOrWhiteSpace(names[index])) names[index] = CreateUnusedTableColumnName(used, index + 1);
                }
            } else if (targetWidth < names.Length) {
                Array.Resize(ref names, targetWidth);
            }
            return ValidateTableColumnNames(names);
        }

        private void RelocateTableTotalsRows(
            (int r1, int c1, int r2, int c2) currentBounds,
            (int r1, int c1, int r2, int c2) targetBounds,
            int totalsRows) {
            if (totalsRows < 1) return;
            int sourceFirstRow = currentBounds.r2 - totalsRows + 1;
            int targetFirstRow = targetBounds.r2 - totalsRows + 1;
            if (sourceFirstRow == targetFirstRow) return;

            int survivingLastColumn = Math.Min(currentBounds.c2, targetBounds.c2);
            var allSourceCells = EnumerateCellsWithEffectiveCoordinates()
                .Where(item => item.Row >= sourceFirstRow
                    && item.Row <= currentBounds.r2
                    && item.Column >= currentBounds.c1
                    && item.Column <= currentBounds.c2)
                .ToList();
            var sourceCells = allSourceCells
                .Where(item => item.Column <= survivingLastColumn)
                .ToList();
            var sourceElements = new HashSet<Cell>(sourceCells.Select(item => item.Cell));
            var discardedElements = new HashSet<Cell>(allSourceCells
                .Where(item => !sourceElements.Contains(item.Cell))
                .Select(item => item.Cell));
            foreach (Cell destination in EnumerateCellsWithEffectiveCoordinates()
                .Where(item => item.Row >= targetFirstRow
                    && item.Row <= targetBounds.r2
                    && item.Column >= targetBounds.c1
                    && item.Column <= targetBounds.c2
                    && !sourceElements.Contains(item.Cell))
                .Select(item => item.Cell)) {
                discardedElements.Add(destination);
            }
            foreach (Cell discarded in discardedElements) {
                ClearCellValueMetadata(discarded);
                discarded.Remove();
            }

            var relocated = sourceCells.Select(item => (
                SourceRow: item.Row,
                SourceColumn: item.Column,
                TargetRow: targetFirstRow + item.Row - sourceFirstRow,
                TargetColumn: targetBounds.c1 + item.Column - currentBounds.c1,
                Cell: (Cell)item.Cell.CloneNode(true))).ToList();
            foreach (var source in sourceCells) source.Cell.Remove();
            foreach (var item in relocated) {
                item.Cell.CellReference = A1.CellReference(item.TargetRow, item.TargetColumn);
                if (item.Cell.CellFormula != null && item.TargetRow != item.SourceRow) {
                    item.Cell.CellFormula.Text = TranslateCopiedFormula(
                        item.Cell.CellFormula.Text,
                        item.SourceRow,
                        item.SourceColumn,
                        item.TargetRow,
                        item.TargetColumn,
                        transpose: false);
                    item.Cell.CellFormula.CalculateCell = true;
                    item.Cell.CellValue = null;
                }
                PutClonedCell(item.TargetRow, item.TargetColumn, item.Cell);
            }
            foreach (Row row in WorksheetRoot.Descendants<Row>().Where(row => !row.Elements<Cell>().Any()).ToList()) {
                row.Remove();
            }
        }

        private void SynchronizeQueryTableSchema(
            Table table,
            TableColumns columns,
            IReadOnlyList<string> names) {
            TableDefinitionPart? tablePart = _worksheetPart.TableDefinitionParts
                .FirstOrDefault(part => ReferenceEquals(part.Table, table));
            if (tablePart == null) return;
            foreach (QueryTablePart queryPart in tablePart.QueryTableParts) {
                QueryTable? queryTable = queryPart.QueryTable;
                if (queryTable == null) continue;
                ExcelDocument.SynchronizeNativeQueryFields(queryTable, columns, names);
                queryTable.Save();
            }
        }

        private static string[] ValidateTableColumnNames(IReadOnlyList<string> columnNames) {
            if (columnNames.Count == 0) throw new ArgumentException("A table must contain at least one column.", nameof(columnNames));
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new string[columnNames.Count];
            for (int index = 0; index < columnNames.Count; index++) {
                string name = columnNames[index]?.Trim() ?? string.Empty;
                if (name.Length == 0) throw new ArgumentException($"Table column {index + 1} has no name.", nameof(columnNames));
                if (name.Length > 32_767) throw new ArgumentException($"Table column {index + 1} exceeds Excel's cell-text limit.", nameof(columnNames));
                if (!used.Add(name)) throw new ArgumentException($"Table column name '{name}' is duplicated.", nameof(columnNames));
                result[index] = name;
            }
            return result;
        }

        private static void RemapTableResizeSortReferences(
            Table table,
            (int r1, int c1, int r2, int c2) currentBounds,
            (int r1, int c1, int r2, int c2) targetBounds,
            int headerRows,
            int totalsRows) {
            int currentDataFirstRow = currentBounds.r1 + headerRows;
            int currentDataLastRow = currentBounds.r2 - totalsRows;
            int targetDataFirstRow = targetBounds.r1 + headerRows;
            int targetDataLastRow = targetBounds.r2 - totalsRows;
            foreach (OpenXmlElement element in table.Descendants().Where(element =>
                string.Equals(element.LocalName, "sortState", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.LocalName, "sortCondition", StringComparison.OrdinalIgnoreCase)).ToList()) {
                OpenXmlAttribute? referenceAttribute = element.GetAttributes()
                    .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "ref", StringComparison.OrdinalIgnoreCase));
                if (referenceAttribute == null
                    || !ExcelReference.TryParse(referenceAttribute.Value.Value, out ExcelReference? reference)) continue;
                reference!.GetBounds(out int r1, out int c1, out int r2, out int c2);
                if (r1 < currentBounds.r1 || c1 < currentBounds.c1
                    || r2 > currentBounds.r2 || c2 > currentBounds.c2) continue;
                bool intersectsTarget = r1 <= targetBounds.r2 && r2 >= targetBounds.r1
                    && c1 <= targetBounds.c2 && c2 >= targetBounds.c1;
                if (!intersectsTarget) {
                    element.Remove();
                    continue;
                }
                int mappedR1 = r1 == currentBounds.r1
                    ? targetBounds.r1
                    : r1 == currentDataFirstRow
                        ? targetDataFirstRow
                        : Math.Max(r1, targetBounds.r1);
                int mappedC1 = c1 == currentBounds.c1 ? targetBounds.c1 : Math.Max(c1, targetBounds.c1);
                int mappedR2 = r2 == currentBounds.r2
                    ? targetBounds.r2
                    : r2 == currentDataLastRow
                        ? targetDataLastRow
                        : Math.Min(r2, targetBounds.r2);
                int mappedC2 = c2 == currentBounds.c2 ? targetBounds.c2 : Math.Min(c2, targetBounds.c2);
                if (mappedR1 > mappedR2 || mappedC1 > mappedC2) {
                    element.Remove();
                    continue;
                }
                element.SetAttribute(new OpenXmlAttribute(
                    referenceAttribute.Value.Prefix,
                    referenceAttribute.Value.LocalName,
                    referenceAttribute.Value.NamespaceUri,
                    reference.WithCoordinates(reference.Kind, mappedR1, mappedC1, mappedR2, mappedC2).ToString()));
            }
            foreach (OpenXmlElement sortState in table.Descendants().Where(element =>
                string.Equals(element.LocalName, "sortState", StringComparison.OrdinalIgnoreCase)).ToList()) {
                if (!sortState.Descendants().Any(element =>
                    string.Equals(element.LocalName, "sortCondition", StringComparison.OrdinalIgnoreCase))) {
                    sortState.Remove();
                }
            }
        }
    }

    public partial class ExcelDocument {
        internal void RewriteTableNameReferences(IEnumerable<string> oldNames, string newName) {
            var aliases = new HashSet<string>(
                oldNames.Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);
            RewriteFormulaRoots(text => ExcelFormulaSyntaxTree.Parse(text).RewriteTableNames(name =>
                aliases.Contains(name) ? newName : name));
            foreach (PivotTableCacheDefinitionPart cachePart in WorkbookPartRoot.PivotTableCacheDefinitionParts) {
                foreach (WorksheetSource source in cachePart.PivotCacheDefinition?.Descendants<WorksheetSource>() ?? Enumerable.Empty<WorksheetSource>()) {
                    if (source.Name?.Value is string sourceName && aliases.Contains(sourceName)) source.Name = newName;
                }
                cachePart.PivotCacheDefinition?.Save();
            }
        }

        internal void RewriteTableColumnReferences(
            string tableName,
            IReadOnlyDictionary<string, string> renames,
            Table owner) {
            RewriteFormulaRoots(text => ExcelFormulaSyntaxTree.Parse(text).RewriteStructuredReferences((name, selector) => {
                if (name != null && string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase)) {
                    return name + ReplaceStructuredColumns(selector, renames);
                }
                return (name ?? string.Empty) + selector;
            }));
            RewriteOwnerWorksheetTableFormulas(owner, text =>
                ExcelFormulaSyntaxTree.Parse(text).RewriteStructuredReferences((name, selector) => name == null
                    ? ReplaceStructuredColumns(selector, renames)
                    : name + selector));
            foreach (OpenXmlLeafTextElement formula in owner.Descendants<OpenXmlLeafTextElement>().Where(IsFormulaLeaf)) {
                formula.Text = ExcelFormulaSyntaxTree.Parse(formula.Text).RewriteStructuredReferences((name, selector) => name == null
                    ? ReplaceStructuredColumns(selector, renames)
                    : name + selector);
            }
        }

        internal void InvalidateTableColumnReferences(string tableName, IReadOnlyCollection<string> removedNames, Table owner) {
            bool ContainsRemoved(string selector) => ExcelFormulaSyntaxTree.ContainsStructuredColumn(selector, removedNames);
            RewriteFormulaRoots(text => ExcelFormulaSyntaxTree.Parse(text).RewriteStructuredReferences((name, selector) =>
                name != null && string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase) && ContainsRemoved(selector)
                    ? null
                    : (name ?? string.Empty) + selector));
            RewriteOwnerWorksheetTableFormulas(owner, text =>
                ExcelFormulaSyntaxTree.Parse(text).RewriteStructuredReferences((name, selector) =>
                    name == null && ContainsRemoved(selector) ? null : (name ?? string.Empty) + selector));
            foreach (OpenXmlLeafTextElement formula in owner.Descendants<OpenXmlLeafTextElement>().Where(IsFormulaLeaf)) {
                formula.Text = ExcelFormulaSyntaxTree.Parse(formula.Text).RewriteStructuredReferences((name, selector) =>
                    name == null && ContainsRemoved(selector) ? null : (name ?? string.Empty) + selector);
            }
        }

        private void RewriteOwnerWorksheetTableFormulas(Table owner, Func<string, string> rewrite) {
            if (!ExcelReference.TryParse(owner.Reference?.Value, out ExcelReference? tableRange)) return;
            WorksheetPart? worksheetPart = WorkbookPartRoot.WorksheetParts.FirstOrDefault(candidate =>
                candidate.TableDefinitionParts.Any(part => ReferenceEquals(part.Table, owner)));
            if (worksheetPart?.Worksheet == null) return;
            foreach (Cell cell in worksheetPart.Worksheet.Descendants<Cell>()) {
                if (cell.CellFormula?.Text is not string formula
                    || !A1.TryParseCellReferenceFast(cell.CellReference?.Value, out int row, out int column)
                    || !tableRange!.Contains(row, column)) {
                    continue;
                }
                cell.CellFormula.Text = rewrite(formula);
            }
        }

        private void RewriteFormulaRoots(Func<string, string> rewrite) {
            RewriteFormulaLeaves(WorkbookPartRoot.Workbook, rewrite);
            var rewrittenChartParts = new HashSet<OpenXmlPart>();
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                RewriteFormulaLeaves(worksheetPart.Worksheet, rewrite);
                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) RewriteFormulaLeaves(tablePart.Table, rewrite);
                RewriteDrawingFormulaRoots(worksheetPart.DrawingsPart, rewrittenChartParts, rewrite);
                foreach (PivotTablePart pivotPart in worksheetPart.PivotTableParts) RewriteFormulaLeaves(pivotPart.PivotTableDefinition, rewrite);
            }
            foreach (ChartsheetPart chartsheetPart in WorkbookPartRoot.ChartsheetParts) {
                RewriteDrawingFormulaRoots(chartsheetPart.DrawingsPart, rewrittenChartParts, rewrite);
            }
        }

        private static void RewriteDrawingFormulaRoots(
            DrawingsPart? drawingsPart,
            HashSet<OpenXmlPart> rewrittenParts,
            Func<string, string> rewrite) {
            if (drawingsPart == null) return;
            foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                if (rewrittenParts.Add(chartPart)) RewriteFormulaLeaves(chartPart.ChartSpace, rewrite);
            }
            foreach (ExtendedChartPart chartPart in drawingsPart.ExtendedChartParts) {
                if (rewrittenParts.Add(chartPart)) RewriteFormulaLeaves(chartPart.ChartSpace, rewrite);
            }
        }

        private static void RewriteFormulaLeaves(OpenXmlPartRootElement? root, Func<string, string> rewrite) {
            if (root == null) return;
            bool changed = false;
            foreach (OpenXmlLeafTextElement leaf in root.Descendants<OpenXmlLeafTextElement>().Where(IsFormulaLeaf)) {
                string rewritten = rewrite(leaf.Text);
                if (string.Equals(rewritten, leaf.Text, StringComparison.Ordinal)) continue;
                leaf.Text = rewritten;
                changed = true;
            }
            if (changed) root.Save();
        }

        private static bool IsFormulaLeaf(OpenXmlLeafTextElement leaf) =>
            string.Equals(leaf.LocalName, "f", StringComparison.OrdinalIgnoreCase)
            || leaf.LocalName.IndexOf("formula", StringComparison.OrdinalIgnoreCase) >= 0
            || string.Equals(leaf.LocalName, "definedName", StringComparison.OrdinalIgnoreCase);

        private static string ReplaceStructuredColumns(
            string selector,
            IReadOnlyDictionary<string, string> renames) =>
            ExcelFormulaSyntaxTree.RewriteStructuredColumns(selector, renames);
    }
}
