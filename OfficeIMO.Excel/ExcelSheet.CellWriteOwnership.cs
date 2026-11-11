using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private int _tableSchemaHeaderWriteDepth;

        // Public table creation always installs a TablePart relationship on the worksheet.
        // Check that small root element before enumerating parts on ordinary cell writes.
        private bool HasTableDefinitions => WorksheetRoot.GetFirstChild<TableParts>()?.HasChildren == true;

        /// <summary>Finds a visible table header without inspecting worksheet cell payloads.</summary>
        private (string Id, TableDefinitionPart Part)[]? _tableDefinitionParts;
        private TableParts? _tableDefinitionPartsList;
        private int _tableDefinitionPartsListCount;

        /// <summary>
        /// Cell writes ask for table headers once per cell. Enumerating the SDK's related parts each
        /// time allocates and walks every relationship, so the parts are kept while the worksheet's
        /// tableParts list is unchanged and every kept relationship id still resolves to the same part.
        /// </summary>
        private (string Id, TableDefinitionPart Part)[] GetTableDefinitionParts() {
            TableParts? list = WorksheetRoot.GetFirstChild<TableParts>();
            int count = 0;
            if (list != null) {
                for (OpenXmlElement? child = list.FirstChild; child != null; child = child.NextSibling()) count++;
            }
            (string Id, TableDefinitionPart Part)[]? cached = _tableDefinitionParts;
            if (cached != null && ReferenceEquals(list, _tableDefinitionPartsList) && count == _tableDefinitionPartsListCount) {
                bool current = true;
                foreach ((string id, TableDefinitionPart part) in cached) {
                    if (!_worksheetPart.TryGetPartById(id, out OpenXmlPart? resolved) || !ReferenceEquals(resolved, part)) {
                        current = false;
                        break;
                    }
                }
                if (current) return cached;
            }

            cached = _worksheetPart.TableDefinitionParts.Select(part => (_worksheetPart.GetIdOfPart(part), part)).ToArray();
            _tableDefinitionParts = cached;
            _tableDefinitionPartsList = list;
            _tableDefinitionPartsListCount = count;
            return cached;
        }

        /// <summary>
        /// Every cell write asks whether it lands on a table header, so the header rows of a worksheet's tables are
        /// indexed by row. The index belongs to the worksheet part and serves while the tableParts list and its last
        /// child are unchanged and no table range, header count or table part has changed since it was built; code
        /// that changes those calls <see cref="ExcelDocument.InvalidateTableLayouts"/>. A remembered header is
        /// confirmed against the live table. Once the package is handed out for direct editing, lookups scan.
        /// </summary>
        private sealed class TableHeaderIndex {
            private readonly Dictionary<int, (Table Table, string Reference, int Left, int Right)[]> _rows = new();

            internal TableHeaderIndex(int epoch, TableParts? list, IEnumerable<TableDefinitionPart> parts) {
                Epoch = epoch;
                List = list;
                LastChild = list?.LastChild;
                foreach (TableDefinitionPart part in parts) {
                    Table? table = part.Table;
                    string? reference = table?.Reference?.Value;
                    if (table == null || reference == null || (table.HeaderRowCount?.Value ?? 1U) == 0U
                        || !ExcelA1.TryParseRange(reference, out int top, out int left, out _, out int right)) continue;
                    var entry = (table, reference, left, right);
                    _rows[top] = _rows.TryGetValue(top, out var existing) ? existing.Concat(new[] { entry }).ToArray() : new[] { entry };
                }
            }

            internal int Epoch { get; }
            internal TableParts? List { get; }
            internal OpenXmlElement? LastChild { get; }

            /// <summary>True or false when the index answers, null when a remembered table changed and it must be rebuilt.</summary>
            internal bool? Find(int row, int column, out Table table, out int headerIndex) {
                table = null!;
                headerIndex = -1;
                if (!_rows.TryGetValue(row, out var entries)) return false;
                foreach (var entry in entries) {
                    if (!string.Equals(entry.Table.Reference?.Value, entry.Reference, StringComparison.Ordinal)
                        || (entry.Table.HeaderRowCount?.Value ?? 1U) == 0U) return null;
                    if (column < entry.Left || column > entry.Right) continue;
                    table = entry.Table;
                    headerIndex = column - entry.Left;
                    return true;
                }
                return false;
            }
        }

        private bool TryGetTableHeader(int row, int column, out Table table, out int headerIndex) {
            table = null!;
            headerIndex = -1;
            if (!HasTableDefinitions) return false;
            if (!_excelDocument.IsOpenXmlDocumentExposed) {
                TableParts? list = WorksheetRoot.GetFirstChild<TableParts>();
                int epoch = _excelDocument.TableLayoutEpoch;
                TableHeaderIndex? index = Persistence.TableHeaders;
                if (index == null || index.Epoch != epoch || !ReferenceEquals(index.List, list)
                    || !ReferenceEquals(index.LastChild, list?.LastChild)) {
                    index = new TableHeaderIndex(epoch, list, GetTableDefinitionParts().Select(entry => entry.Part));
                    Persistence.TableHeaders = index;
                }
                bool? found = index.Find(row, column, out table, out headerIndex);
                if (found.HasValue) return found.Value;
                Persistence.TableHeaders = null;
            }
            foreach ((_, TableDefinitionPart part) in GetTableDefinitionParts()) {
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
            foreach ((_, TableDefinitionPart part) in GetTableDefinitionParts()) {
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

        /// <summary>Protects schema labels when removing a formula array cache.</summary>
        private void EnsureTableHeaderValuesNotCleared(int top, int left, int bottom, int right) {
            if (RangeContainsTableHeader(top, left, bottom, right))
                throw new ArgumentException("An array cache cannot erase table column labels. Replace the table header with a literal name first.");
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
