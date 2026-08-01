using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void RemoveQueryBackedTableBinding(string tableName, bool preserveTable) {
            ApplyTransactionalMutation(_ => {
                TableDefinitionPart tablePart = _worksheetPart.TableDefinitionParts.FirstOrDefault(part =>
                    string.Equals(part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException($"Query-backed table '{tableName}' was not found on worksheet '{Name}'.");
                QueryTablePart queryPart = tablePart.QueryTableParts.FirstOrDefault()
                    ?? throw new InvalidOperationException($"Table '{tableName}' is not query-backed.");
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
                    tablePart.Table?.Save();
                }
                WorksheetRoot.Save();
                return 1;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
        }

        internal string ReplaceQueryBackedTableData(
            string tableName,
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
                var current = A1.ParseRange(table.Reference?.Value
                    ?? throw new InvalidDataException("Query-backed table range is missing."));
                int targetLastColumn = checked(current.c1 + columnNames.Count - 1);
                int targetLastRow = checked(current.r1 + rows.Count);
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

                RemoveCellsInRange(current.r1, current.c1, current.r2, current.c2);
                for (int column = 0; column < columnNames.Count; column++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    CellValueCoreNoMaterialize(current.r1, current.c1 + column, columnNames[column]);
                }
                for (int row = 0; row < rows.Count; row++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    object?[] values = rows[row];
                    for (int column = 0; column < values.Length; column++) {
                        CellValueCoreNoMaterialize(current.r1 + row + 1, current.c1 + column, values[column]);
                    }
                }

                table.Reference = updatedRange;
                AutoFilter? filter = table.GetFirstChild<AutoFilter>();
                if (filter != null) {
                    filter.Reference = updatedRange;
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
                return checked(rows.Count * columnNames.Count + columnNames.Count);
            }, new ExcelMutationPlanOptions(), cancellationToken);
            return updatedRange!;
        }

        private void EnsureQueryRefreshExpansionIsEmpty(ExcelReference current, ExcelReference target) {
            bool conflict = WorksheetRoot.Descendants<Cell>().Any(cell =>
                TryGetCellCoordinates(cell, out int row, out int column)
                && target.Contains(row, column)
                && !current.Contains(row, column)
                && (cell.CellFormula != null || cell.CellValue != null || cell.InlineString != null));
            if (conflict) {
                throw new InvalidOperationException(
                    "Query refresh expansion would overwrite populated worksheet cells outside the current table range.");
            }
        }
    }
}
