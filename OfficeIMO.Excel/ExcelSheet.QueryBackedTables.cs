using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void RemoveQueryBackedTableBinding(
            string tableName,
            uint expectedConnectionId,
            bool preserveTable) {
            ApplyTransactionalMutation(_ => {
                TableDefinitionPart tablePart = _worksheetPart.TableDefinitionParts.FirstOrDefault(part =>
                    string.Equals(part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException($"Query-backed table '{tableName}' was not found on worksheet '{Name}'.");
                QueryTablePart queryPart = tablePart.QueryTableParts.FirstOrDefault()
                    ?? throw new InvalidOperationException($"Table '{tableName}' is not query-backed.");
                if (queryPart.QueryTable?.ConnectionId?.Value != expectedConnectionId) {
                    throw new InvalidOperationException(
                        $"Query-backed table '{tableName}' changed before its binding could be removed.");
                }
                tablePart.DeletePart(queryPart);
                if (!preserveTable) {
                    string relationshipId = _worksheetPart.GetIdOfPart(tablePart);
                    TableParts? tableParts = WorksheetRoot.GetFirstChild<TableParts>();
                    tableParts?.Elements<TablePart>().FirstOrDefault(item => item.Id?.Value == relationshipId)?.Remove();
                    if (tableParts != null) {
                        if (tableParts.Elements<TablePart>().Any()) tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();
                        else tableParts.Remove();
                    }
                    DeleteTableDefinitionPart(tablePart);
                } else {
                    Table table = tablePart.Table
                        ?? throw new InvalidDataException($"Query-backed table '{tableName}' has no table definition.");
                    table.ConnectionId = null;
                    foreach (TableColumn column in table.Descendants<TableColumn>()) {
                        column.QueryTableFieldId = null;
                    }
                    table.Save();
                }
                WorksheetRoot.Save();
                return 1;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
        }

        internal string ReplaceQueryBackedTableData(
            string tableName,
            uint expectedConnectionId,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<object?[]> rows,
            CancellationToken cancellationToken) {
            string? updatedRange = null;
            ApplyTransactionalMutation(_ => {
                TableDefinitionPart tablePart = _worksheetPart.TableDefinitionParts.FirstOrDefault(part =>
                    string.Equals(part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException($"Query-backed table '{tableName}' was not found on worksheet '{Name}'.");
                Table table = tablePart.Table ?? throw new InvalidDataException("Query-backed table definition is missing.");
                QueryTablePart queryPart = tablePart.QueryTableParts.FirstOrDefault()
                    ?? throw new InvalidOperationException($"Table '{tableName}' is not query-backed.");
                if (queryPart.QueryTable?.ConnectionId?.Value != expectedConnectionId) {
                    throw new InvalidOperationException(
                        $"Query-backed table '{tableName}' changed while its query was executing.");
                }
                var current = A1.ParseRange(table.Reference?.Value
                    ?? throw new InvalidDataException("Query-backed table range is missing."));
                NormalizeImplicitCellReferences();
                int headerRowCount = checked((int)(table.HeaderRowCount?.Value ?? 1U));
                if (headerRowCount < 0 || headerRowCount > 1) {
                    throw new InvalidDataException("Query-backed table header-row metadata is unsupported.");
                }
                int totalsRowCount = checked((int)(table.TotalsRowCount?.Value
                    ?? (table.TotalsRowShown?.Value == true ? 1U : 0U)));
                int currentRowCount = current.r2 - current.r1 + 1;
                if (totalsRowCount < 0 || headerRowCount + totalsRowCount > currentRowCount) {
                    throw new InvalidDataException("Query-backed table row metadata is inconsistent with its range.");
                }
                int currentTotalsStartRow = current.r2 - totalsRowCount + 1;
                int targetLastColumn = checked(current.c1 + columnNames.Count - 1);
                int targetRowCount = checked(headerRowCount + rows.Count + totalsRowCount);
                if (targetRowCount == 0) {
                    throw new InvalidOperationException("A headerless query-backed table cannot represent an empty result.");
                }
                int targetLastRow = checked(current.r1 + targetRowCount - 1);
                if (targetLastColumn > 16_384 || targetLastRow > 1_048_576) {
                    throw new InvalidOperationException("Query result exceeds worksheet capacity.");
                }
                updatedRange = A1.CellReference(current.r1, current.c1) + ":"
                    + A1.CellReference(targetLastRow, targetLastColumn);
                ExcelReference currentReference = ExcelReference.Parse(table.Reference!.Value!);
                ExcelReference targetReference = ExcelReference.Parse(updatedRange);
                EnsureQueryRefreshExpansionIsEmpty(currentReference, targetReference);
                EnsureNoIntersectingOwnedStructures(
                    targetReference,
                    "Query refresh would overlap another table, merged cells, an array or data-table formula, or PivotTable output.",
                    excludedTable: table);

                TableColumns columns = table.TableColumns ??= new TableColumns();
                List<TableColumn> existing = columns.Elements<TableColumn>().ToList();
                string[] removedNames = existing.Skip(columnNames.Count)
                    .Select(column => column.Name?.Value ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .ToArray();
                var renames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                uint nextId = existing.Count == 0 ? 1U : existing.Max(column => column.Id?.Value ?? 0U) + 1U;
                for (int index = 0; index < columnNames.Count; index++) {
                    TableColumn column;
                    if (index < existing.Count) {
                        column = existing[index];
                        string oldName = column.Name?.Value ?? string.Empty;
                        if (oldName.Length > 0 && !string.Equals(oldName, columnNames[index], StringComparison.Ordinal)) {
                            renames[oldName] = columnNames[index];
                        }
                    } else {
                        column = columns.AppendChild(new TableColumn { Id = nextId++ });
                    }
                    column.Name = columnNames[index];
                }
                for (int index = existing.Count - 1; index >= columnNames.Count; index--) existing[index].Remove();
                columns.Count = (uint)columnNames.Count;

                List<(int RowOffset, int ColumnOffset, Cell Cell)> totalsCells =
                    ExtractQueryRefreshTotalsCells(current, totalsRowCount, targetLastColumn);
                if (headerRowCount > 0) {
                    for (int column = 0; column < columnNames.Count; column++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        CellValueCoreNoMaterialize(current.r1, current.c1 + column, columnNames[column]);
                    }
                }
                for (int row = 0; row < rows.Count; row++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    object?[] values = rows[row];
                    for (int column = 0; column < values.Length; column++) {
                        CellValueCoreNoMaterialize(current.r1 + headerRowCount + row, current.c1 + column, values[column]);
                    }
                }
                int targetTotalsStartRow = targetLastRow - totalsRowCount + 1;
                foreach ((int rowOffset, int columnOffset, Cell clone) in totalsCells) {
                    if (columnOffset >= columnNames.Count) continue;
                    int sourceRow = currentTotalsStartRow + rowOffset;
                    int sourceColumn = current.c1 + columnOffset;
                    int targetRow = targetTotalsStartRow + rowOffset;
                    int targetColumn = current.c1 + columnOffset;
                    clone.CellReference = A1.CellReference(targetRow, targetColumn);
                    if (clone.CellFormula != null && targetRow != sourceRow) {
                        clone.CellFormula.Text = TranslateCopiedFormula(
                            clone.CellFormula.Text,
                            sourceRow,
                            sourceColumn,
                            targetRow,
                            targetColumn,
                            transpose: false);
                        clone.CellFormula.CalculateCell = true;
                        clone.CellValue = null;
                    }
                    PutClonedCell(targetRow, targetColumn, clone);
                }

                table.Reference = updatedRange;
                AutoFilter? filter = table.GetFirstChild<AutoFilter>();
                if (filter != null && headerRowCount == 0) {
                    filter.Remove();
                } else if (filter != null) {
                    int filterLastRow = targetLastRow - totalsRowCount;
                    filter.Reference = A1.CellReference(current.r1, current.c1) + ":"
                        + A1.CellReference(filterLastRow, targetLastColumn);
                    foreach (FilterColumn stale in filter.Elements<FilterColumn>()
                        .Where(item => (item.ColumnId?.Value ?? uint.MaxValue) >= (uint)columnNames.Count).ToList()) stale.Remove();
                }
                string stableName = table.Name?.Value ?? table.DisplayName?.Value ?? tableName;
                if (removedNames.Length > 0) _excelDocument.InvalidateTableColumnReferences(stableName, removedNames, table);
                if (renames.Count > 0) _excelDocument.RewriteTableColumnReferences(stableName, renames, table);
                ExcelDocument.SynchronizeNativeQueryFields(queryPart.QueryTable!, columns, columnNames);
                queryPart.QueryTable!.Save();
                table.Save();
                WorksheetRoot.Save();
                _excelDocument.CleanupCalculationArtifacts(
                    save: false,
                    ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
                return checked(rows.Count * columnNames.Count
                    + headerRowCount * columnNames.Count
                    + totalsCells.Count);
            }, new ExcelMutationPlanOptions(), cancellationToken);
            return updatedRange!;
        }

        private List<(int RowOffset, int ColumnOffset, Cell Cell)> ExtractQueryRefreshTotalsCells(
            (int r1, int c1, int r2, int c2) current,
            int totalsRowCount,
            int targetLastColumn) {
            int totalsStartRow = current.r2 - totalsRowCount + 1;
            int preservedLastColumn = Math.Min(current.c2, targetLastColumn);
            var currentCells = EnumerateCellsWithEffectiveCoordinates()
                .Where(item => item.Row >= current.r1
                    && item.Row <= current.r2
                    && item.Column >= current.c1
                    && item.Column <= current.c2)
                .ToList();
            var preserved = totalsRowCount == 0
                ? new List<(Cell Cell, int Row, int Column)>()
                : currentCells.Where(item => item.Row >= totalsStartRow
                    && item.Column <= preservedLastColumn).ToList();
            var preservedElements = new HashSet<Cell>(preserved.Select(item => item.Cell));

            foreach (var discarded in currentCells.Where(item => !preservedElements.Contains(item.Cell))) {
                ClearCellValueMetadata(discarded.Cell);
                discarded.Cell.Remove();
            }
            var snapshots = preserved.Select(item => (
                item.Row - totalsStartRow,
                item.Column - current.c1,
                (Cell)item.Cell.CloneNode(true))).ToList();
            foreach (var item in preserved) item.Cell.Remove();
            foreach (Row row in WorksheetRoot.Descendants<Row>().Where(row => !row.Elements<Cell>().Any()).ToList()) {
                row.Remove();
            }
            return snapshots;
        }

        private void EnsureQueryRefreshExpansionIsEmpty(ExcelReference current, ExcelReference target) {
            bool conflict = EnumerateCellsWithEffectiveCoordinates().Any(item =>
                target.Contains(item.Row, item.Column)
                && !current.Contains(item.Row, item.Column)
                && (item.Cell.CellFormula != null || item.Cell.CellValue != null || item.Cell.InlineString != null));
            if (conflict) {
                throw new InvalidOperationException(
                    "Query refresh expansion would overwrite populated worksheet cells outside the current table range.");
            }
        }
    }
}
