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

                if (CanStartUnqualifiedStructuredReference(formula, index)) {
                    string reference = formula.Substring(index, end - index);
                    if (TryGetUnqualifiedCurrentRowSections(reference, out List<string> sections)
                        && TryResolveTableReferenceRange(
                            table,
                            sections,
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

        private static bool TryGetUnqualifiedCurrentRowSections(string reference, out List<string> sections) {
            if (!TryParseStructuredTableReference("T" + reference, out _, out sections)) {
                return false;
            }

            if (sections.Count == 1) {
                string section = sections[0];
                if ((section.Length > 1 && section[0] == '@')
                    || string.Equals(section, "#This Row", StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
                if (IsStructuredTableAreaSpecifier(section)) {
                    return false;
                }

                sections = new List<string> { "#This Row", section };
                return true;
            }

            return sections.Count == 2
                && string.Equals(sections[0], "#This Row", StringComparison.OrdinalIgnoreCase);
        }
    }
}
