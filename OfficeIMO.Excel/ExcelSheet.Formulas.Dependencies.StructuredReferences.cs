using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaDependencyTableRange {
            internal FormulaDependencyTableRange(Table table, int firstRow, int firstColumn, int lastRow, int lastColumn) {
                Table = table;
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
            }

            internal Table Table { get; }
            internal int FirstRow { get; }
            internal int FirstColumn { get; }
            internal int LastRow { get; }
            internal int LastColumn { get; }
        }

        private sealed class FormulaDependencyTableCatalog {
            private readonly List<FormulaDependencyTableRange> _ranges;
            private readonly int[] _prefixMaximumLastRows;

            internal FormulaDependencyTableCatalog(IEnumerable<Table> tables) {
                _ranges = tables
                    .Select(TryCreateRange)
                    .Where(range => range.HasValue)
                    .Select(range => range!.Value)
                    .OrderBy(range => range.FirstRow)
                    .ThenBy(range => range.FirstColumn)
                    .ToList();
                _prefixMaximumLastRows = new int[_ranges.Count];
                int maximumLastRow = 0;
                for (int index = 0; index < _ranges.Count; index++) {
                    maximumLastRow = Math.Max(maximumLastRow, _ranges[index].LastRow);
                    _prefixMaximumLastRows[index] = maximumLastRow;
                }
            }

            internal bool TryGetContainingTable(int row, int column, out Table table) {
                int low = 0;
                int high = _ranges.Count - 1;
                int candidate = -1;
                while (low <= high) {
                    int middle = low + ((high - low) / 2);
                    if (_ranges[middle].FirstRow <= row) {
                        candidate = middle;
                        low = middle + 1;
                    } else {
                        high = middle - 1;
                    }
                }

                for (int index = candidate; index >= 0 && _prefixMaximumLastRows[index] >= row; index--) {
                    FormulaDependencyTableRange range = _ranges[index];
                    if (range.LastRow >= row
                        && range.FirstColumn <= column
                        && range.LastColumn >= column) {
                        table = range.Table;
                        return true;
                    }
                }

                table = null!;
                return false;
            }

            private static FormulaDependencyTableRange? TryCreateRange(Table table) {
                if (table.Reference?.Value == null
                    || !A1.TryParseRange(
                        table.Reference.Value.Replace("$", string.Empty),
                        out int firstRow,
                        out int firstColumn,
                        out int lastRow,
                        out int lastColumn)) {
                    return null;
                }

                return new FormulaDependencyTableRange(table, firstRow, firstColumn, lastRow, lastColumn);
            }
        }

        private FormulaDependencyTableCatalog GetFormulaDependencyTables() {
            return new FormulaDependencyTableCatalog(
                _worksheetPart.TableDefinitionParts
                    .Select(part => part.Table)
                    .OfType<Table>());
        }

        private void AddUnqualifiedCurrentRowDependencyMatches(
            string formula,
            int sourceRow,
            int sourceColumn,
            FormulaDependencyTableCatalog tables,
            ICollection<FormulaDependencyReferenceMatch> dependencyMatches) {
            if (formula.IndexOf('[') < 0
                || !tables.TryGetContainingTable(sourceRow, sourceColumn, out Table table)) {
                return;
            }

            for (int index = 0; index < formula.Length; index++) {
                if (formula[index] != '[') {
                    continue;
                }
                if (!TryFindStructuredReferenceEnd(formula, index, out int end)) {
                    break;
                }

                if (CanStartUnqualifiedStructuredReference(formula, index)
                    && !ContinuesAsExternalWorkbookQualifier(formula, end)) {
                    string reference = formula.Substring(index, end - index);
                    if (TryGetUnqualifiedCurrentRowReference(reference, out FormulaStructuredTableReference structuredReference)
                        && TryResolveTableReferenceRange(
                            table,
                            structuredReference,
                            sourceRow,
                            out int firstRow,
                            out int firstColumn,
                            out int lastRow,
                            out int lastColumn)) {
                        string start = A1.CellReference(firstRow, firstColumn);
                        string endReference = A1.CellReference(lastRow, lastColumn);
                        dependencyMatches.Add(new FormulaDependencyReferenceMatch(
                            index,
                            reference.Length,
                            firstRow == lastRow && firstColumn == lastColumn
                                ? start
                                : start + ":" + endReference));
                    }
                }

                index = end - 1;
            }
        }

        private static bool CanStartUnqualifiedStructuredReference(string formula, int index) {
            if (index == 0) {
                return true;
            }

            char previous = formula[index - 1];
            return !IsFormulaAliasIdentifierCharacter(previous)
                && previous != '['
                && previous != ']'
                && previous != '!'
                && previous != '\'';
        }

        private static bool ContinuesAsExternalWorkbookQualifier(string formula, int end) {
            return end < formula.Length && IsFormulaAliasIdentifierCharacter(formula[end]);
        }

        private static bool TryFindStructuredReferenceEnd(string formula, int start, out int end) {
            int depth = 0;
            for (int index = start; index < formula.Length; index++) {
                if (formula[index] == '[') {
                    depth++;
                } else if (formula[index] == ']') {
                    depth--;
                    if (depth == 0) {
                        end = index + 1;
                        return true;
                    }
                }
            }

            end = start;
            return false;
        }

        private static bool TryGetUnqualifiedCurrentRowReference(
            string reference,
            out FormulaStructuredTableReference structuredReference) {
            if (!TryParseStructuredTableReference("T" + reference, out _, out structuredReference)) {
                return false;
            }

            if (structuredReference.AreaIsExplicit) {
                return string.Equals(structuredReference.Area, "#This Row", StringComparison.OrdinalIgnoreCase);
            }

            structuredReference = structuredReference.WithArea("#This Row");
            return true;
        }

        private bool TryResolveUnqualifiedCurrentRowTableReferenceRange(
            string token,
            int? currentRow,
            out ExcelSheet sheet,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn) {
            sheet = this;
            firstRow = 0;
            firstColumn = 0;
            lastRow = 0;
            lastColumn = 0;
            if (!currentRow.HasValue
                || _formulaEvaluationCellReference == null
                || !TryParseCellReference(_formulaEvaluationCellReference, out int evaluationRow, out int evaluationColumn)
                || evaluationRow != currentRow.Value) {
                return false;
            }

            return TryResolveUnqualifiedCurrentRowTableReferenceRange(
                token,
                currentRow,
                evaluationColumn,
                out sheet,
                out firstRow,
                out firstColumn,
                out lastRow,
                out lastColumn);
        }

        private bool TryResolveUnqualifiedCurrentRowTableReferenceRange(
            string token,
            int? currentRow,
            int? currentColumn,
            out ExcelSheet sheet,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn) {
            sheet = this;
            firstRow = 0;
            firstColumn = 0;
            lastRow = 0;
            lastColumn = 0;
            if (!currentRow.HasValue
                || !currentColumn.HasValue
                || !TryGetUnqualifiedCurrentRowReference(token, out FormulaStructuredTableReference structuredReference)
                || !TryGetContainingFormulaTable(currentRow.Value, currentColumn.Value, out Table table)) {
                return false;
            }

            return TryResolveTableReferenceRange(
                table,
                structuredReference,
                currentRow,
                out firstRow,
                out firstColumn,
                out lastRow,
                out lastColumn);
        }

        private bool TryGetContainingFormulaTable(int row, int column, out Table table) {
            foreach (Table candidate in _worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .OfType<Table>()) {
                if (candidate.Reference?.Value != null
                    && A1.TryParseRange(
                        candidate.Reference.Value.Replace("$", string.Empty),
                        out int firstRow,
                        out int firstColumn,
                        out int lastRow,
                        out int lastColumn)
                    && row >= firstRow
                    && row <= lastRow
                    && column >= firstColumn
                    && column <= lastColumn) {
                    table = candidate;
                    return true;
                }
            }

            table = null!;
            return false;
        }
    }
}
