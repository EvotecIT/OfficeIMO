using System.Xml;
using System.Security.Cryptography;
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

            // One pass over the relationships; asking the SDK for each part's id walks them again per table.
            var parts = new List<(string Id, TableDefinitionPart Part)>();
            foreach (IdPartPair pair in _worksheetPart.Parts) {
                if (pair.OpenXmlPart is TableDefinitionPart tablePart) parts.Add((pair.RelationshipId, tablePart));
            }
            cached = parts.ToArray();
            _tableDefinitionParts = cached;
            _tableDefinitionPartsList = list;
            _tableDefinitionPartsListCount = count;
            return cached;
        }

        /// <summary>
        /// Every cell write asks whether it lands on a table header, and every new table checks the others for overlap,
        /// so the tables of a worksheet are indexed: header rows by row, every range, the tablePart ids and the widest
        /// column. The index belongs to the worksheet part and serves while the tableParts list and its last child are
        /// unchanged and no table range, header count or table part has changed since it was built; code that changes
        /// those calls <see cref="ExcelDocument.InvalidateTableLayouts"/>, and a new table is appended to it. A remembered
        /// table is confirmed against the live one. Once the package is handed out for direct editing, lookups scan.
        /// </summary>
        private sealed class TableHeaderIndex {
            private readonly Dictionary<int, (Table Table, string Reference, int Left, int Right)[]> _rows = new();
            private readonly List<TableLayoutEntry> _tables = new();
            private readonly HashSet<string> _tablePartIds = new(StringComparer.Ordinal);
            private int _lastColumn;

            internal TableHeaderIndex(int epoch, TableParts? list, IEnumerable<TableDefinitionPart> parts) {
                Epoch = epoch;
                List = list;
                LastChild = list?.LastChild;
                foreach (TableDefinitionPart part in parts) {
                    if (part.Table is Table table) Add(table);
                }
                if (list != null) {
                    foreach (TablePart tablePart in list.Elements<TablePart>()) {
                        if (tablePart.Id?.Value is string id) _tablePartIds.Add(id);
                    }
                }
            }

            internal int Epoch { get; private set; }
            internal TableParts? List { get; private set; }
            internal OpenXmlElement? LastChild { get; private set; }

            /// <summary>Takes a table that was just added under the write lock, so the index stays current.</summary>
            internal void Append(Table table, string? tablePartId, int epoch, TableParts list) {
                Add(table);
                if (tablePartId != null) _tablePartIds.Add(tablePartId);
                Epoch = epoch;
                List = list;
                LastChild = list.LastChild;
            }

            internal bool ContainsTablePart(string id) => _tablePartIds.Contains(id);

            private void Add(Table table) {
                string? reference = table.Reference?.Value;
                if (reference == null) return;
                bool strict = ExcelA1.TryParseStrictRange(reference, out int strictTop, out int strictLeft, out int strictBottom, out int strictRight);
                bool parsed = ExcelA1.TryParseRange(reference, out int top, out int left, out _, out int right);
                _tables.Add(new TableLayoutEntry(table, reference, strict, strictTop, strictLeft, strictBottom, strictRight));
                if (parsed && right > _lastColumn) _lastColumn = right;
                if (!parsed || (table.HeaderRowCount?.Value ?? 1U) == 0U) return;
                var entry = (table, reference, left, right);
                _rows[top] = _rows.TryGetValue(top, out var existing) ? existing.Concat(new[] { entry }).ToArray() : new[] { entry };
            }

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

            /// <summary>
            /// Whether a strictly parsed table range overlaps the given cells; null when the overlapping table changed.
            /// Range edits retire the index through the layout epoch, so only a hit is confirmed against the live table.
            /// </summary>
            internal bool? Overlaps(int top, int left, int bottom, int right) {
                foreach (TableLayoutEntry entry in _tables) {
                    if (entry.Strict && left <= entry.StrictRight && right >= entry.StrictLeft
                        && top <= entry.StrictBottom && bottom >= entry.StrictTop) return entry.IsCurrent ? true : null;
                }
                return false;
            }

            /// <summary>The right-most column any table reaches.</summary>
            internal int LastColumn => _lastColumn;

            private readonly struct TableLayoutEntry {
                internal TableLayoutEntry(Table table, string reference, bool strict, int strictTop, int strictLeft, int strictBottom, int strictRight) {
                    Table = table;
                    Reference = reference;
                    Strict = strict;
                    StrictTop = strictTop;
                    StrictLeft = strictLeft;
                    StrictBottom = strictBottom;
                    StrictRight = strictRight;
                }

                internal Table Table { get; }
                internal string Reference { get; }
                internal bool Strict { get; }
                internal int StrictTop { get; }
                internal int StrictLeft { get; }
                internal int StrictBottom { get; }
                internal int StrictRight { get; }
                internal bool IsCurrent => string.Equals(Table.Reference?.Value, Reference, StringComparison.Ordinal);
            }
        }

        /// <summary>Gets the worksheet's table index, rebuilding it when tables changed; null once the package is handed out.</summary>
        private TableHeaderIndex? GetCurrentTableLayout() {
            if (_excelDocument.IsOpenXmlDocumentExposed) return null;
            TableParts? list = WorksheetRoot.GetFirstChild<TableParts>();
            int epoch = _excelDocument.TableLayoutEpoch;
            TableHeaderIndex? index = Persistence.TableHeaders;
            if (index == null || index.Epoch != epoch || !ReferenceEquals(index.List, list)
                || !ReferenceEquals(index.LastChild, list?.LastChild)) {
                index = new TableHeaderIndex(epoch, list, GetTableDefinitionParts().Select(entry => entry.Part));
                Persistence.TableHeaders = index;
            }
            return index;
        }

        private bool TryGetTableHeader(int row, int column, out Table table, out int headerIndex) {
            table = null!;
            headerIndex = -1;
            if (!HasTableDefinitions) return false;
            TableHeaderIndex? index = GetCurrentTableLayout();
            if (index != null) {
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
            _excelDocument.EffectiveDateTimeOffsetWriteStrategy,
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

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<
            WorksheetPart, DynamicSpillOwnership> DynamicSpillOwnerships =
            new System.Runtime.CompilerServices.ConditionalWeakTable<
                WorksheetPart, DynamicSpillOwnership>();

        private sealed class DynamicSpillOwnership {
            internal readonly Dictionary<long, DynamicSpillCacheSnapshot> Cells =
                new Dictionary<long, DynamicSpillCacheSnapshot>();
            internal bool WriteIndexInitialized;
            internal readonly List<(int Top, int Left, int Bottom, int Right)> WriteRanges =
                new List<(int Top, int Left, int Bottom, int Right)>();
            internal bool OriginalFingerprintAttempted;
            internal byte[]? OriginalFingerprint;
            internal bool OriginalScanAttempted;
            internal readonly HashSet<long> WrittenOwners = new HashSet<long>();
        }

        private DynamicSpillOwnership? _spillOwnership;

        private DynamicSpillOwnership SpillOwnership => _spillOwnership ??= DynamicSpillOwnerships.GetOrCreateValue(_worksheetPart);

        private Dictionary<long, DynamicSpillCacheSnapshot> DynamicSpillCacheSnapshots =>
            SpillOwnership.Cells;

        private sealed class DynamicSpillCacheSnapshot {
            internal Cell? SourceCell;
            internal CellValues? Type;
            internal string? Value;
            internal uint? ValueMetaIndex;
        }

        private static DynamicSpillCacheSnapshot SnapshotDynamicSpillCell(Cell cell) => new DynamicSpillCacheSnapshot {
            SourceCell = cell,
            Type = cell.DataType?.Value,
            Value = ExcelWorksheetScan.FindValue(cell)?.Text,
            ValueMetaIndex = cell.ValueMetaIndex?.Value
        };

        private static bool MatchesDynamicSpillSnapshot(Cell cell, DynamicSpillCacheSnapshot snapshot) =>
            cell.DataType?.Value == snapshot.Type
            && string.Equals(ExcelWorksheetScan.FindValue(cell)?.Text, snapshot.Value, StringComparison.Ordinal)
            && cell.ValueMetaIndex?.Value == snapshot.ValueMetaIndex;

        private void InvalidateDynamicArrayWriteIndex() => SpillOwnership.WriteIndexInitialized = false;

        private void EnsureDynamicArrayWriteIndex() {
            DynamicSpillOwnership ownership = SpillOwnership;
            if (ownership.WriteIndexInitialized) return;
            ownership.WriteRanges.Clear();
            if (_excelDocument.WorkbookPartRoot.CellMetadataPart != null) {
                ExcelCellMetadataIndex? metadata = _excelDocument.FormulaWorkbookModel.CellMetadata;
                foreach (Cell cell in ExcelWorksheetScan.FormulaCells(WorksheetRoot)) {
                    if (ExcelWorksheetScan.FindFormula(cell)?.FormulaType?.Value != CellFormulaValues.Array
                        || cell.CellMetaIndex == null
                        || ExcelCellMetadataIndex.ArrayInfo(WrapFormulaCell(cell), metadata)?.IsDynamic != true
                        || !TryFixedArrayBounds(ExcelWorksheetScan.FindFormula(cell)!.Reference?.Value ?? "",
                            out int top, out int left, out int bottom, out int right)) continue;
                    ownership.WriteRanges.Add((top, left, bottom, right));
                }
            }
            ownership.WriteIndexInitialized = true;
        }

        private void EnsureDynamicArrayCellWritable(int row, int column) {
            EnsureDynamicArrayWriteIndex();
            List<(int Top, int Left, int Bottom, int Right)> ranges = SpillOwnership.WriteRanges;
            for (int i = 0; i < ranges.Count; i++) {
                var range = ranges[i];
                if (row >= range.Top && row <= range.Bottom && column >= range.Left && column <= range.Right)
                    throw new InvalidOperationException($"Cell '{ExcelA1.CellReference(row, column)}' belongs to a dynamic array; clear its array formula first.");
            }
        }

        internal Cell GetWritableValueCell(int row, int column) {
            EnsureCellPayloadWritable(row, column);
            return GetCell(row, column);
        }

        private void EnsureDynamicArrayRangeWritable(int top, int left, int bottom, int right) {
            EnsureDynamicArrayWriteIndex();
            if (SpillOwnership.WriteRanges.Any(range =>
                RangesOverlapInclusive((top, left, bottom, right), range)))
                throw new InvalidOperationException("The range overlaps a dynamic array; clear its array formula first.");
        }

        private void WriteDynamicArrayFormulaCache(DynamicArrayPlan plan) {
            Cell anchor = SdkCell(plan.Owner.Cell);
            SpillOwnership.WrittenOwners.Add(DynamicCellKey(plan.Owner.Top, plan.Owner.Left));
            string? previousReference = ExcelWorksheetScan.FindFormula(anchor)?.Reference?.Value;
            if (plan.ErrorSubtype != 0) {
                ExcelWorksheetScan.FindFormula(anchor)!.Reference = anchor.CellReference!.Value!;
                TryWriteRichFormulaError(anchor, "#SPILL!", plan);
            } else {
                FormulaArrayValue array = plan.Array!;
                ExcelWorksheetScan.FindFormula(anchor)!.Reference = ExcelA1.CellReference(plan.Top, plan.Left) + ":"
                    + ExcelA1.CellReference(plan.Bottom, plan.Right);
                for (int row = plan.Top; row <= plan.Bottom; row++)
                    for (int column = plan.Left; column <= plan.Right; column++) {
                        Cell cell = row == plan.Top && column == plan.Left ? anchor : GetCell(row, column);
                        cell.InlineString = null;
                        cell.ValueMetaIndex = null;
                        SetFormulaCachedValue(WrapFormulaCell(cell), array.Values[(row - plan.Top) * array.Columns + column - plan.Left]);
                        if (cell != anchor)
                            DynamicSpillCacheSnapshots[DynamicCellKey(row, column)] = SnapshotDynamicSpillCell(cell);
                    }
            }
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = plan.Owner.Top; row <= plan.Owner.Bottom; row++)
                for (int column = plan.Owner.Left; column <= plan.Owner.Right; column++) {
                    if (row == plan.Owner.Top && column == plan.Owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    if (!index.Cells.TryGetValue(DynamicCellKey(row, column), out ExcelFormulaCell handle)) continue;
                    Cell cell = SdkCell(handle);
                    long key = DynamicCellKey(row, column);
                    if (!IsOwnedDynamicSpillCachedCell(cell, key)) continue;
                    cell.CellValue = null;
                    cell.InlineString = null;
                    cell.DataType = null;
                    cell.ValueMetaIndex = null;
                    DynamicSpillCacheSnapshots.Remove(key);
                    if (cell.StyleIndex == null && !cell.HasChildren) cell.Remove();
                }
            if (!string.Equals(previousReference, ExcelWorksheetScan.FindFormula(anchor)?.Reference?.Value, StringComparison.Ordinal))
                InvalidateDynamicArrayWriteIndex();
        }

        private bool IsDynamicSpillObstacle(FixedArrayOwner owner, ExcelFormulaCell handle, int row, int column,
            FormulaArgumentValue expected) {
            Cell cell = SdkCell(handle);
            if (row >= owner.Top && row <= owner.Bottom && column >= owner.Left && column <= owner.Right)
                return !IsOwnedDynamicSpillCachedCell(cell, DynamicCellKey(row, column), expected);
            return ExcelWorksheetScan.FindValue(cell) != null || cell.InlineString != null || cell.ValueMetaIndex != null
                || ExcelWorksheetScan.FindFormula(cell) != null || cell.CellMetaIndex != null;
        }

        private bool IsOwnedDynamicSpillCachedCell(ExcelFormulaCell handle, long key, FormulaArgumentValue? expected = null) =>
            IsOwnedDynamicSpillCachedCell(SdkCell(handle), key, expected);

        private bool IsOwnedDynamicSpillCachedCell(Cell cell, long key, FormulaArgumentValue? expected = null) {
            if (ExcelWorksheetScan.FindFormula(cell) != null || cell.CellMetaIndex != null || cell.InlineString != null) return false;
            if (DynamicSpillCacheSnapshots.TryGetValue(key, out DynamicSpillCacheSnapshot? snapshot))
                return MatchesDynamicSpillSnapshot(cell, snapshot);
            if (ExcelWorksheetScan.FindValue(cell) == null && cell.ValueMetaIndex == null) return true;
            if (expected == null) return false;
            FormulaArgumentValue value = expected.Value;
            if (value.IsError)
                return cell.DataType?.Value == CellValues.Error
                    && ResolveRichValueError(cell, ExcelWorksheetScan.FindValue(cell)?.Text) == value.ErrorCode;
            if (cell.ValueMetaIndex != null) return false;
            if (value.IsBoolean)
                return cell.DataType?.Value == CellValues.Boolean
                    && ExcelWorksheetScan.FindValue(cell)?.Text == (value.Number == 0 ? "0" : "1");
            if (value.Number.HasValue)
                return (cell.DataType == null
                        || cell.DataType.Value == CellValues.Number)
                    && OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(ExcelWorksheetScan.FindValue(cell)?.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double numeric)
                    && numeric == value.Number.Value;
            return value.Text != null && string.Equals(GetCellText(cell), value.Text, StringComparison.Ordinal);
        }

    }
}

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const long MaxOriginalDynamicSpillXmlBytes = 64L * 1024 * 1024;
        private const string SpreadsheetMainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string StrictSpreadsheetMainNamespace = "http://purl.oclc.org/ooxml/spreadsheetml/main";

        private void CaptureOriginalDynamicSpillFingerprintIfSafe() {
            // Workbook-only edits leave this part untouched. Once the raw package
            // is exposed or its root loaded, prior external edits are ambiguous.
            if (!_excelDocument.IsOpenXmlDocumentExposed && !_worksheetPart.IsRootElementLoaded)
                CaptureOriginalDynamicSpillFingerprint();
        }

        // A fingerprint costs no retained copy of a loaded worksheet. It also prevents a
        // later worksheet-part Save from becoming an accidental ownership baseline.
        private void CaptureOriginalDynamicSpillFingerprint() {
            DynamicSpillOwnership ownership = SpillOwnership;
            lock (ownership) {
                if (ownership.OriginalFingerprintAttempted) return;
                ownership.OriginalFingerprintAttempted = true;
                try {
                    if (!HasDynamicArrayMetadataType()) return;

                    ownership.OriginalFingerprint = ReadBoundedWorksheetFingerprint();
                } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                    || ex is InvalidDataException || ex is XmlException) {
                    // The normal formula path still validates metadata; this optional baseline fails closed.
                }
            }
        }

        private bool HasDynamicArrayMetadataType() {
            var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
            if (part == null) return false;
            if (!part.IsRootElementLoaded) ValidateInCellImageMetadataPart(part, "Cell metadata");
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                MaxCharactersInDocument = MaximumRichValueMetadataBytes
            });
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "metadataType"
                    && IsSpreadsheetMainNamespace(reader.NamespaceURI)
                    && string.Equals(reader.GetAttribute("name"), "XLDAPR", StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private byte[]? ReadBoundedWorksheetFingerprint() {
            try {
                using Stream stream = _worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > MaxOriginalDynamicSpillXmlBytes) return null;
                using SHA256 sha = SHA256.Create();
                return sha.ComputeHash(stream);
            } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                || ex is NotSupportedException) {
                return null;
            }
        }

        private sealed class OriginalDynamicOwnerCaches {
            internal readonly FixedArrayOwner Owner;
            internal readonly Dictionary<long, DynamicSpillCacheSnapshot> Cells =
                new Dictionary<long, DynamicSpillCacheSnapshot>();
            internal readonly HashSet<long> SeenChildren = new HashSet<long>();
            internal bool AnchorSeen;
            internal bool Valid = true;

            internal OriginalDynamicOwnerCaches(FixedArrayOwner owner) => Owner = owner;
        }

        private void TryCaptureOriginalDynamicSpillCaches(FixedArraySheetIndex? sheetIndex = null) {
            DynamicSpillOwnership ownership = SpillOwnership;
            if (ownership.OriginalScanAttempted) return;
            ownership.OriginalScanAttempted = true;
            if (ownership.OriginalFingerprint == null) return;

            byte[]? currentFingerprint = ReadBoundedWorksheetFingerprint();
            if (currentFingerprint == null || !currentFingerprint.SequenceEqual(ownership.OriginalFingerprint))
                return;

            sheetIndex ??= GetFixedArraySheetIndex();

            var anchors = new Dictionary<long, OriginalDynamicOwnerCaches>();
            var children = new Dictionary<long, OriginalDynamicOwnerCaches>();
            var states = new List<OriginalDynamicOwnerCaches>();
            long remainingCells = MaxResolvedFormulaRangeCells;
            foreach (FixedArrayOwner owner in sheetIndex.DynamicOwners) {
                long count = (long)(owner.Bottom - owner.Top + 1) * (owner.Right - owner.Left + 1);
                if (ownership.WrittenOwners.Contains(DynamicCellKey(owner.Top, owner.Left))
                    || count < 1 || count > remainingCells) continue;
                remainingCells -= count;
                var state = new OriginalDynamicOwnerCaches(owner);
                states.Add(state);
                long anchorKey = DynamicCellKey(owner.Top, owner.Left);
                if (anchors.ContainsKey(anchorKey) || children.ContainsKey(anchorKey)) return;
                anchors.Add(anchorKey, state);
                for (int row = owner.Top; row <= owner.Bottom; row++)
                    for (int column = owner.Left; column <= owner.Right; column++) {
                        long key = DynamicCellKey(row, column);
                        if (key == anchorKey) continue;
                        if (children.ContainsKey(key) || anchors.ContainsKey(key)) return;
                        children.Add(key, state);
                    }
            }
            if (states.Count == 0 || !ReadOriginalDynamicSpillCells(anchors, children)) return;
            foreach (OriginalDynamicOwnerCaches state in states)
                if (state.Valid && state.AnchorSeen)
                    foreach (var entry in state.Cells)
                        if (!ownership.Cells.ContainsKey(entry.Key)) {
                            if (sheetIndex.Cells.TryGetValue(entry.Key, out ExcelFormulaCell source))
                                entry.Value.SourceCell = (source.Record as SdkFormulaCellRecord)?.Cell;
                            ownership.Cells.Add(entry.Key, entry.Value);
                        }
        }

        private bool ReadOriginalDynamicSpillCells(
            Dictionary<long, OriginalDynamicOwnerCaches> anchors,
            Dictionary<long, OriginalDynamicOwnerCaches> children) {
            try {
                using Stream stream = _worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > MaxOriginalDynamicSpillXmlBytes) return false;
                using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    MaxCharactersInDocument = MaxOriginalDynamicSpillXmlBytes
                });
                int sheetDataDepth = -1;
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == sheetDataDepth
                        && reader.LocalName == "sheetData") sheetDataDepth = -1;
                    if (reader.NodeType != XmlNodeType.Element || !IsSpreadsheetMainNamespace(reader.NamespaceURI))
                        continue;
                    if (reader.LocalName == "sheetData") {
                        sheetDataDepth = reader.Depth;
                        continue;
                    }
                    if (sheetDataDepth < 0 || reader.LocalName != "c") continue;
                    if (!TryParseCellReference(reader.GetAttribute("r") ?? "", out int row, out int column))
                        continue;
                    long key = DynamicCellKey(row, column);
                    if (anchors.TryGetValue(key, out OriginalDynamicOwnerCaches? anchor)) {
                        if (anchor.AnchorSeen) anchor.Valid = false;
                        anchor.AnchorSeen = true;
                        anchor.Valid &= OriginalDynamicAnchorMatches(reader, anchor.Owner);
                        continue;
                    }
                    if (!children.TryGetValue(key, out OriginalDynamicOwnerCaches? child)) continue;
                    if (!child.SeenChildren.Add(key)) {
                        child.Valid = false;
                        continue;
                    }
                    if (TryReadOriginalPlainCache(reader, out DynamicSpillCacheSnapshot? snapshot))
                        child.Cells.Add(key, snapshot!);
                }
                return true;
            } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                || ex is XmlException || ex is NotSupportedException) {
                return false;
            }
        }

        private static bool OriginalDynamicAnchorMatches(XmlReader reader, FixedArrayOwner owner) {
            string? cellMetadata = reader.GetAttribute("cm");
            if (!uint.TryParse(cellMetadata, out uint metadata)
                || metadata != owner.Cell.CellMetaIndex) return false;
            string expectedRange = ExcelA1.CellReference(owner.Top, owner.Left) + ":"
                + ExcelA1.CellReference(owner.Bottom, owner.Right);
            using XmlReader child = reader.ReadSubtree();
            while (child.Read()) {
                if (child.NodeType != XmlNodeType.Element || child.LocalName != "f"
                    || !IsSpreadsheetMainNamespace(child.NamespaceURI)) continue;
                string? type = child.GetAttribute("t");
                string? range = child.GetAttribute("ref");
                string formula = child.ReadElementContentAsString();
                return type == "array" && string.Equals(range, expectedRange, StringComparison.Ordinal)
                    && string.Equals(formula, owner.Cell.FormulaText, StringComparison.Ordinal);
            }
            return false;
        }

        private static bool TryReadOriginalPlainCache(XmlReader reader, out DynamicSpillCacheSnapshot? snapshot) {
            snapshot = null;
            if (reader.GetAttribute("cm") != null) return false;
            string? type = reader.GetAttribute("t");
            CellValues? parsedType = type switch {
                null => null,
                "b" => CellValues.Boolean,
                "d" => CellValues.Date,
                "e" => CellValues.Error,
                "n" => CellValues.Number,
                "s" => CellValues.SharedString,
                "str" => CellValues.String,
                _ => null
            };
            if (type != null && parsedType == null) return false;
            uint? valueMetadata = null;
            if (reader.GetAttribute("vm") is string vm) {
                if (!uint.TryParse(vm, out uint parsed)) return false;
                valueMetadata = parsed;
            }

            string? value = null;
            bool hasValue = false;
            using XmlReader child = reader.ReadSubtree();
            while (child.Read()) {
                if (child.NodeType != XmlNodeType.Element || child.Depth != 1) continue;
                if (child.LocalName != "v" || hasValue
                    || !IsSpreadsheetMainNamespace(child.NamespaceURI)) return false;
                hasValue = true;
                value = child.ReadElementContentAsString();
            }
            if (!hasValue && valueMetadata == null) return false;
            snapshot = new DynamicSpillCacheSnapshot {
                Type = parsedType,
                Value = value,
                ValueMetaIndex = valueMetadata
            };
            return true;
        }

        private static bool IsSpreadsheetMainNamespace(string namespaceUri) =>
            string.Equals(namespaceUri, SpreadsheetMainNamespace, StringComparison.Ordinal)
            || string.Equals(namespaceUri, StrictSpreadsheetMainNamespace, StringComparison.Ordinal);
    }
}
