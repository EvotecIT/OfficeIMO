using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private int _tableSchemaHeaderWriteDepth;

        // Public table creation always installs a TablePart relationship on the worksheet.
        // Check that small root element before enumerating parts on ordinary cell writes.
        private bool HasTableDefinitions => WorksheetRoot.GetFirstChild<TableParts>()?.HasChildren == true;

        /// <summary>Finds a visible table header without inspecting worksheet cell payloads.</summary>
        private bool TryGetTableHeader(int row, int column, out Table table, out int headerIndex) {
            table = null!;
            headerIndex = -1;
            if (!HasTableDefinitions) return false;
            foreach (var part in _worksheetPart.TableDefinitionParts) {
                Table? candidate = part.Table;
                if (candidate == null || (candidate.HeaderRowCount?.Value ?? 1U) == 0U
                    || !ExcelA1.TryParseRange(candidate.Reference?.Value ?? string.Empty,
                        out int top, out int left, out _, out int right)
                    || row != top || column < left || column > right) continue;
                table = candidate;
                headerIndex = column - left;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Routes a header through the schema owner; its own physical writes do not reenter it.
        /// </summary>
        private bool TrySetTableHeaderValue(int row, int column, object? value) {
            if (_tableSchemaHeaderWriteDepth > 0) return false;
            if (!TryGetTableHeader(row, column, out Table table, out int index)) return false;
            SetTableHeaderValue(table, index, value);
            return true;
        }

        /// <summary>Writes a literal header whose schema owner already validated its name and planned range.</summary>
        private void WriteOwnedTableSchemaHeader(int row, int column, string name) {
            _tableSchemaHeaderWriteDepth++;
            try { CellValueCoreNoMaterialize(row, column, name); }
            finally { _tableSchemaHeaderWriteDepth--; }
        }

        private string SetTableHeaderValue(Table table, int index, object? value) {
            EnsureTableHeaderWritable(table, index);
            string[] names = table.TableColumns?.Elements<TableColumn>()
                .Select(column => column.Name?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>();
            if (index >= names.Length) throw new InvalidOperationException("The table header has no matching schema column.");
            names[index] = GetTableHeaderValueText(value);
            return SetTableSchema(table.Name?.Value ?? table.Reference!.Value!, names)[index].Name;
        }

        /// <summary>Uses the cell coercer's invariant text, as table creation does, without adding shared strings.</summary>
        private string GetTableHeaderValueText(object? value) => CoerceValueHelper.Coerce(
            value,
            text => new CellValue(text),
            _excelDocument.DateTimeOffsetWriteStrategy,
            _excelDocument.DateSystem).cellValue.Text;

        private void EnsureTableHeaderWritable(Table table, int index) {
            var bounds = ExcelA1.ParseRange(table.Reference!.Value!);
            EnsureDynamicArrayCellWritable(bounds.r1, bounds.c1 + index);
        }

        /// <summary>Non-schema payloads cannot replace the literal names that own table columns.</summary>
        private void EnsureCellPayloadWritable(int row, int column) {
            EnsureDynamicArrayCellWritable(row, column);
            if (_tableSchemaHeaderWriteDepth == 0 && !_materializingPendingCellValueDirectSaveBuffer
                && !_excelDocument.IsMaterializingDeferredDataSetImport
                && TryGetTableHeader(row, column, out _, out _))
                throw new InvalidOperationException("Table headers require literal column names. Use SetValue, SetRichText or SetTableSchema to change a header.");
        }

        private bool RangeContainsTableHeader(int top, int left, int bottom, int right) {
            if (!HasTableDefinitions) return false;
            foreach (var part in _worksheetPart.TableDefinitionParts) {
                Table? table = part.Table;
                if (table == null || (table.HeaderRowCount?.Value ?? 1U) == 0U
                    || !ExcelA1.TryParseRange(table.Reference?.Value ?? string.Empty,
                        out int headerRow, out int firstColumn, out _, out int lastColumn)) continue;
                if (headerRow >= top && headerRow <= bottom && firstColumn <= right && lastColumn >= left) return true;
            }
            return false;
        }

        private void EnsureTableHeaderFormulaRangeWritable(int top, int left, int bottom, int right) {
            if (RangeContainsTableHeader(top, left, bottom, right))
                throw new InvalidOperationException("An array formula cannot replace literal table-column headers.");
        }

        private bool FormulaTargetsTableHeader(Cell cell) {
            if (!HasTableDefinitions) return false;
            bool mayBeHeader = true;
            if (string.IsNullOrEmpty(cell.CellReference?.Value) && cell.Parent is Row parent
                && parent.RowIndex?.Value is uint nativeRow && nativeRow > 0U && nativeRow <= ExcelA1.MaxRows) {
                mayBeHeader = RangeContainsTableHeader((int)nativeRow, 1, (int)nativeRow, ExcelA1.MaxColumns);
            }
            if (mayBeHeader && TryGetEffectiveCellCoordinates(cell, out int row, out int column)
                && TryGetTableHeader(row, column, out _, out _)) return true;
            CellFormula? formula = ExcelWorksheetScan.FindFormula(cell);
            return formula?.FormulaType?.Value == CellFormulaValues.Array
                && TryFixedArrayBounds(formula.Reference?.Value ?? string.Empty,
                    out int top, out int left, out int bottom, out int right)
                && RangeContainsTableHeader(top, left, bottom, right);
        }

        private void EnsureTableHeaderValuesNotCleared(int top, int left, int bottom, int right) {
            if (RangeContainsTableHeader(top, left, bottom, right))
                throw new ArgumentException("Table column names cannot be empty. Clear data rows or use SetTableSchema to rename headers.");
        }

        /// <summary>Renames the schema while retaining rich-run properties and canonical trimming.</summary>
        private void PrepareTableHeaderRichText(int row, int column, InlineString inline) {
            if (!TryGetTableHeader(row, column, out Table table, out int index)) return;
            Text[] nodes = EnumerateRichTextNodes(inline).ToArray();
            string[] segments = nodes.Select(node => Utilities.ExcelSanitizer.DecodeXString(node.Text ?? string.Empty)).ToArray();
            string text = string.Concat(segments);
            string name = SetTableHeaderValue(table, index, text);
            if (string.Equals(text, name, StringComparison.Ordinal)) return;
            int first = text.Length - text.TrimStart().Length;
            int last = first + name.Length;
            int offset = 0;
            for (int part = 0; part < nodes.Length; part++) {
                string segment = segments[part];
                int start = Math.Min(segment.Length, Math.Max(0, first - offset));
                int end = Math.Min(segment.Length, Math.Max(start, last - offset));
                nodes[part].Text = Utilities.ExcelSanitizer.EncodeXString(segment.Substring(start, end - start));
                offset += segment.Length;
            }
        }

        /// <summary>
        /// Resolves a bulk write's final header names together, so swaps use one canonical
        /// schema rewrite. The prepared writer then receives those same normalized text values.
        /// </summary>
        private IReadOnlyList<(int Row, int Column, object Value)> PrepareTableHeaderValueWrites(
            IReadOnlyList<(int Row, int Column, object Value)> source,
            CancellationToken cancellationToken) {
            if (!HasTableDefinitions) return source;
            return Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                (int Row, int Column, object Value)[]? normalized = null;
                var schemas = new List<(Table Table, string[] Names)>();
                foreach (var part in _worksheetPart.TableDefinitionParts) {
                    Table? table = part.Table;
                    if (table == null || (table.HeaderRowCount?.Value ?? 1U) == 0U
                        || !ExcelA1.TryParseRange(table.Reference?.Value ?? string.Empty,
                            out int top, out int left, out _, out int right)) continue;
                    string[]? names = null;
                    for (int index = 0; index < source.Count; index++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = source[index];
                        if (item.Row != top || item.Column < left || item.Column > right) continue;
                        EnsureDynamicArrayCellWritable(item.Row, item.Column);
                        names ??= table.TableColumns?.Elements<TableColumn>()
                            .Select(column => column.Name?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>();
                        int headerIndex = item.Column - left;
                        if (headerIndex >= names.Length) throw new InvalidOperationException("The table header has no matching schema column.");
                        names[headerIndex] = GetTableHeaderValueText(item.Value);
                    }
                    if (names == null) continue;
                    names = ValidateTableColumnNames(names);
                    normalized ??= source.ToArray();
                    for (int index = 0; index < source.Count; index++) {
                        var item = source[index];
                        if (item.Row == top && item.Column >= left && item.Column <= right)
                            normalized[index] = (item.Row, item.Column, names[item.Column - left]);
                    }
                    schemas.Add((table, names));
                }
                foreach (var schema in schemas) {
                    cancellationToken.ThrowIfCancellationRequested();
                    SetTableSchema(schema.Table.Name?.Value ?? schema.Table.Reference!.Value!, schema.Names);
                }
                return (IReadOnlyList<(int Row, int Column, object Value)>?)normalized ?? source;
            });
        }
    }
}
