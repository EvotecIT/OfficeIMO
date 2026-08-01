using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Builds a bounded dry-run plan for inserting complete worksheet columns.</summary>
        public ExcelStructuralMutationPlan PlanInsertColumns(int firstColumn, int count = 1, ExcelMutationPlanOptions? options = null) =>
            PlanColumnMutation(firstColumn, count, deleting: false, options);

        /// <summary>Builds a bounded dry-run plan for deleting complete worksheet columns.</summary>
        public ExcelStructuralMutationPlan PlanDeleteColumns(int firstColumn, int count = 1, ExcelMutationPlanOptions? options = null) =>
            PlanColumnMutation(firstColumn, count, deleting: true, options);

        /// <summary>Transactionally inserts complete columns and returns post-edit package diagnostics.</summary>
        public ExcelMutationResult InsertColumns(int firstColumn, int count = 1, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanInsertColumns(firstColumn, count, options).Apply(cancellationToken);

        /// <summary>Transactionally deletes complete columns and returns post-edit package diagnostics.</summary>
        public ExcelMutationResult DeleteColumns(int firstColumn, int count = 1, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanDeleteColumns(firstColumn, count, options).Apply(cancellationToken);

        private ExcelStructuralMutationPlan PlanColumnMutation(
            int firstColumn,
            int count,
            bool deleting,
            ExcelMutationPlanOptions? options) {
            ValidateStructuralColumnArguments(firstColumn, count);
            ExcelMutationPlanOptions effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                MutationPlanScanBudget budget = CreateMutationPlanScanBudget(effective);
                int last = firstColumn + count - 1;
                PreflightColumnMutation(firstColumn, last, count, deleting, budget);
                int cells = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget).Count(cell => {
                    int column = cell.CellReference?.Value is string reference ? GetColumnIndex(reference) : 0;
                    return deleting ? column >= firstColumn : column >= firstColumn;
                });
                if (cells > effective.MaximumAffectedCells) {
                    throw new InvalidOperationException($"Column mutation affects {cells} cells, exceeding MaximumAffectedCells ({effective.MaximumAffectedCells}).");
                }
                var impacts = new List<ExcelMutationImpact>();
                if (cells > 0) impacts.Add(new ExcelMutationImpact("cells", cells, "Worksheet cells will be shifted or removed."));
                int formulas = InspectMutationPlanElements(WorksheetRoot.Descendants<CellFormula>(), budget)
                    .Count(formula => ExcelFormulaSyntaxTree.Parse(formula.Text ?? string.Empty)
                        .Nodes.OfType<ExcelFormulaReferenceSyntax>().Any());
                if (formulas > 0) impacts.Add(new ExcelMutationImpact("formulas", formulas, "Parsed formula references may be rewritten."));
                int tables = InspectMutationPlanElements(_worksheetPart.TableDefinitionParts, budget).Count();
                if (tables > 0) impacts.Add(new ExcelMutationImpact("tables", tables, "Intersecting table ranges and schemas will be remapped."));
                int drawings = InspectMutationPlanElements(
                    _worksheetPart.DrawingsPart?.WorksheetDrawing?.ChildElements
                        ?? Enumerable.Empty<OpenXmlElement>(),
                    budget).Count();
                if (drawings > 0) impacts.Add(new ExcelMutationImpact("drawings", drawings, "Cell-anchored drawing columns will be remapped."));
                string range = A1.ColumnIndexToLetters(firstColumn) + ":" + A1.ColumnIndexToLetters(last);
                return new ExcelStructuralMutationPlan(
                    this,
                    deleting ? ExcelStructuralMutationKind.DeleteColumns : ExcelStructuralMutationKind.InsertColumns,
                    range,
                    null,
                    cells,
                    impacts,
                    effective,
                    cancellationToken => {
                        cancellationToken.ThrowIfCancellationRequested();
                        ExcelStructuralMutationPlan current = PlanColumnMutation(firstColumn, count, deleting, effective);
                        ApplyColumnMutation(firstColumn, count, deleting, cancellationToken);
                        return current.AffectedCells;
                    });
            });
        }

        private void PreflightColumnMutation(
            int firstColumn,
            int lastColumn,
            int count,
            bool deleting,
            MutationPlanScanBudget? budget = null) {
            ValidateA1MutationReferenceMode("Structural column edits");
            ValidateWorkbookSharedFormulasForStructuralEdit();
            ValidateStructuralVmlControlSafety();
            ValidateColumnConnectionParameters(firstColumn, count, deleting);
            if (!deleting) ValidateColumnCommentVmlAnchorCapacity(firstColumn, count);
            if (!deleting) {
                int maximumUsedColumn = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget)
                    .Select(cell => cell.CellReference?.Value is string reference ? GetColumnIndex(reference) : 0)
                    .DefaultIfEmpty(0).Max();
                if (maximumUsedColumn >= firstColumn && (long)maximumUsedColumn + count > A1.MaxColumns) {
                    throw new InvalidOperationException("Column insertion would move worksheet content beyond the Excel column limit.");
                }
            }
            foreach (CellFormula formula in InspectMutationPlanElements(WorksheetRoot.Descendants<CellFormula>(), budget).Where(item =>
                item.FormulaType?.Value == CellFormulaValues.Array || item.FormulaType?.Value == CellFormulaValues.DataTable)) {
                if (!ExcelReference.TryParse(formula.Reference?.Value, out ExcelReference? range)) continue;
                range!.GetBounds(out _, out int c1, out _, out int c2);
                bool conflict = deleting
                    ? c1 <= lastColumn && c2 >= firstColumn
                    : c1 < firstColumn && c2 >= firstColumn;
                if (conflict) throw new InvalidOperationException("Column mutation would split an array formula or data table.");
            }
            foreach (TableDefinitionPart part in InspectMutationPlanElements(_worksheetPart.TableDefinitionParts, budget)) {
                Table? table = part.Table;
                if (!deleting || !ExcelReference.TryParse(table?.Reference?.Value, out ExcelReference? range)) continue;
                range!.GetBounds(out _, out int c1, out _, out int c2);
                int overlap = Math.Max(0, Math.Min(c2, lastColumn) - Math.Max(c1, firstColumn) + 1);
                if (overlap >= c2 - c1 + 1) throw new InvalidOperationException($"Column deletion would remove every column from table '{table!.Name?.Value}'.");
            }
            foreach (PivotTablePart part in InspectMutationPlanElements(_worksheetPart.PivotTableParts, budget)) {
                if (!ExcelReference.TryParse(part.PivotTableDefinition?.Location?.Reference?.Value, out ExcelReference? range)) continue;
                range!.GetBounds(out _, out int c1, out _, out int c2);
                bool conflict = deleting
                    ? c1 <= lastColumn && c2 >= firstColumn
                    : c1 < firstColumn && c2 >= firstColumn;
                if (conflict) throw new InvalidOperationException("Column mutation would intersect or split a PivotTable output range.");
            }
        }

        private void ApplyColumnMutation(int firstColumn, int count, bool deleting, CancellationToken cancellationToken) {
            int lastColumn = firstColumn + count - 1;
            MaterializeWorkbookSharedFormulasForStructuralEdit();
            IReadOnlyList<(int Row, int Column, string Name)> pendingHeaders = AdjustTableSchemasForColumnMutation(firstColumn, lastColumn, count, deleting);
            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (Row row in sheetData.Elements<Row>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    List<Cell> cells = row.Elements<Cell>().ToList();
                    if (deleting) {
                        foreach (Cell cell in cells) {
                            int column = cell.CellReference?.Value is string reference ? GetColumnIndex(reference) : 0;
                            if (column >= firstColumn && column <= lastColumn) cell.Remove();
                            else if (column > lastColumn) cell.CellReference = A1.CellReference((int)(row.RowIndex?.Value ?? 0U), column - count);
                        }
                    } else {
                        foreach (Cell cell in cells.OrderByDescending(cell =>
                            cell.CellReference?.Value is string reference ? GetColumnIndex(reference) : 0)) {
                            int column = cell.CellReference?.Value is string reference ? GetColumnIndex(reference) : 0;
                            if (column >= firstColumn) cell.CellReference = A1.CellReference((int)(row.RowIndex?.Value ?? 0U), checked(column + count));
                        }
                    }
                    if (!row.Elements<Cell>().Any()) row.Remove();
                }
            }
            foreach ((int row, int column, string name) in pendingHeaders) CellValueCoreNoMaterialize(row, column, name);
            RewriteColumnDefinitions(firstColumn, lastColumn, count, deleting);
            RemapColumnConnectionParameters(firstColumn, count, deleting, cancellationToken);
            _excelDocument.RewriteColumnMutationReferences(this, firstColumn, count, deleting);
            RemapColumnCommentVml(firstColumn, count, deleting);
            _excelDocument.CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
            ResetMutationCaches();
        }

        private IReadOnlyList<(int Row, int Column, string Name)> AdjustTableSchemasForColumnMutation(int firstColumn, int lastColumn, int count, bool deleting) {
            var pendingHeaders = new List<(int Row, int Column, string Name)>();
            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts) {
                Table? table = part.Table;
                if (!ExcelReference.TryParse(table?.Reference?.Value, out ExcelReference? range)) continue;
                range!.GetBounds(out int r1, out int c1, out _, out int c2);
                TableColumns? columns = table!.TableColumns;
                if (columns == null) continue;
                List<TableColumn> existing = columns.Elements<TableColumn>().ToList();
                if (!deleting && firstColumn > c1 && firstColumn <= c2) {
                    int offset = firstColumn - c1;
                    uint nextId = existing.Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty().Max() + 1U;
                    var used = new HashSet<string>(existing.Select(item => item.Name?.Value ?? string.Empty), StringComparer.OrdinalIgnoreCase);
                    for (int index = 0; index < count; index++) {
                        string name = CreateUnusedTableColumnName(used, offset + index + 1);
                        var added = new TableColumn { Id = nextId++, Name = name };
                        TableColumn? before = columns.Elements<TableColumn>().ElementAtOrDefault(offset + index);
                        if (before == null) columns.Append(added); else columns.InsertBefore(added, before);
                        pendingHeaders.Add((r1, firstColumn + index, name));
                    }
                    AutoFilter? filter = table.GetFirstChild<AutoFilter>();
                    if (filter != null) {
                        foreach (FilterColumn column in filter.Elements<FilterColumn>()) {
                            uint id = column.ColumnId?.Value ?? uint.MaxValue;
                            if (id >= (uint)offset) column.ColumnId = id + (uint)count;
                        }
                    }
                } else if (deleting && firstColumn <= c2 && lastColumn >= c1) {
                    int removeStart = Math.Max(firstColumn, c1) - c1;
                    int removeEnd = Math.Min(lastColumn, c2) - c1;
                    string[] removedNames = existing.Skip(removeStart).Take(removeEnd - removeStart + 1)
                        .Select(item => item.Name?.Value ?? string.Empty).Where(name => name.Length > 0).ToArray();
                    for (int index = removeEnd; index >= removeStart; index--) existing[index].Remove();
                    AutoFilter? filter = table.GetFirstChild<AutoFilter>();
                    if (filter != null) {
                        int removed = removeEnd - removeStart + 1;
                        foreach (FilterColumn column in filter.Elements<FilterColumn>().ToList()) {
                            uint id = column.ColumnId?.Value ?? uint.MaxValue;
                            if (id >= (uint)removeStart && id <= (uint)removeEnd) column.Remove();
                            else if (id > (uint)removeEnd) column.ColumnId = id - (uint)removed;
                        }
                    }
                    string tableName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                    if (tableName.Length > 0 && removedNames.Length > 0) {
                        _excelDocument.InvalidateTableColumnReferences(tableName, removedNames, table);
                    }
                }
                columns.Count = (uint)columns.Elements<TableColumn>().Count();
            }
            return pendingHeaders;
        }

        private static string CreateUnusedTableColumnName(HashSet<string> used, int suggestedIndex) {
            int index = Math.Max(1, suggestedIndex);
            string name;
            do name = "Column" + index++.ToString(CultureInfo.InvariantCulture); while (!used.Add(name));
            return name;
        }

        private void RewriteColumnDefinitions(int firstColumn, int lastColumn, int count, bool deleting) {
            foreach (Columns columns in WorksheetRoot.Elements<Columns>().ToList()) {
                foreach (Column column in columns.Elements<Column>().ToList()) {
                    uint min = column.Min?.Value ?? 1U;
                    uint max = column.Max?.Value ?? min;
                    if (!deleting) {
                        if (min >= (uint)firstColumn) {
                            column.Min = min + (uint)count;
                            column.Max = max + (uint)count;
                        }
                        else if (max >= (uint)firstColumn) column.Max = max + (uint)count;
                    } else if (max < (uint)firstColumn) {
                        continue;
                    } else if (min > (uint)lastColumn) {
                        column.Min = min - (uint)count;
                        column.Max = max - (uint)count;
                    } else {
                        uint survivorsLeft = min < (uint)firstColumn ? (uint)firstColumn - min : 0U;
                        uint survivorsRight = max > (uint)lastColumn ? max - (uint)lastColumn : 0U;
                        if (survivorsLeft == 0U && survivorsRight == 0U) column.Remove();
                        else {
                            column.Min = survivorsLeft > 0U ? min : (uint)firstColumn;
                            column.Max = survivorsRight > 0U ? max - (uint)count : (uint)firstColumn - 1U;
                        }
                    }
                }
                if (!columns.Elements<Column>().Any()) columns.Remove();
            }
        }

        private static void ValidateStructuralColumnArguments(int firstColumn, int count) {
            if (firstColumn < 1 || firstColumn > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(firstColumn));
            if (count < 1 || (long)firstColumn + count - 1L > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(count));
        }
    }
}
