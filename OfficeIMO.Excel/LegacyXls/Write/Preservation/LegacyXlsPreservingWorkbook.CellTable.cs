using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.Utilities;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    internal byte[] WriteEdits(ExcelDocument document, IReadOnlyList<LegacyXlsCellEdit> edits,
        byte[]? filePass, long maximumBytes, CancellationToken token, Func<byte[], byte[]>? transform = null, IReadOnlyList<ExcelLegacySheetBindings.Binding>? bindings = null, bool calculationChanged = false, ExcelNativeCalculationPolicy calculation = default) {
        var structure = bindings == null ? null : new LegacyXlsSheetStructure(bindings, _nativeSheets);
        var compound = PrepareVbaHosts(structure, maximumBytes, token);
        var encodingRecords = structure?.PrepareGlobals(_records) ?? _records;
        var encoding = LegacyXlsWriter.EncodePreservedEdits(document, _model, encodingRecords, structure?.CurrentIndexes ?? _nativeSheets,
            edits, _stringCount, token, structure);
        var layout = BuildEditingLayout(document, encoding, token, structure, calculation);
        if (calculationChanged) ApplyCalculationSettings(document, layout, structure, token);
        if (structure != null) ApplyStructure(layout, structure, encoding, token);
        byte[] workbook = RewriteRecords(new(), encoding.Strings, filePass, maximumBytes, token, layout);
        if (transform != null) workbook = transform(workbook);
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) { [_streamName] = workbook };
        if (transform != null) Compound.LegacyXlsCompoundEncryption.AddEncryptedAuxiliaryStreams(compound, replacements, transform);
        return OfficeCompoundFileWriter.Rewrite(compound, replacements, maxOutputBytes: maximumBytes, cancellationToken: token);
    }

    private EditingLayout BuildEditingLayout(ExcelDocument document, LegacyXlsEditEncoding encoding, CancellationToken token, LegacyXlsSheetStructure? structure = null, ExcelNativeCalculationPolicy calculation = default) {
        var layout = new EditingLayout();
        int globalsEnd = _records.ToList().FindIndex(record => record.Type == 0x000a);
        void AppendAfter(ushort type, IEnumerable<byte[]> payloads) {
            int last = -1;
            for (int i = 0; i < globalsEnd; i++) if (_records[i].Type == type) last = i;
            int before = last < 0 ? _sstIndex >= 0 ? _sstIndex : globalsEnd : last + 1;
            foreach (var payload in payloads) layout.InsertBefore(before, layout.New(type, payload));
        }
        if (encoding.Palette != null) {
            int palette = _records.ToList().FindIndex(record => record.Type == 0x0092);
            if (palette >= 0) layout.Replace[palette] = new[] { new OutputRecord(palette, 0x0092, encoding.Palette) };
            else layout.InsertBefore(globalsEnd, layout.New(0x0092, encoding.Palette));
        }
        AppendAfter(0x0031, encoding.Fonts); AppendAfter(0x041e, encoding.Formats); AppendAfter(0x00e0, encoding.Styles);
        int externIndex = -1;
        for (int i = 0; i < globalsEnd; i++) if (_records[i].Type == 0x0017) externIndex = i;
        int externPosition = externIndex < 0 ? globalsEnd : externIndex;
        foreach (var link in encoding.Externs.SupportingLinkRecords)
            layout.InsertBefore(externPosition, layout.New(link.RecordType, link.Payload));
        if (externIndex >= 0) layout.Replace[externIndex] = new[] { new OutputRecord(externIndex, 0x0017, encoding.Externs.Payload) };
        else layout.InsertBefore(externPosition, layout.New(0x0017, encoding.Externs.Payload));

        foreach (var sheet in _model.Worksheets.Where(sheet => structure == null || structure.CurrentName(sheet.Name) != null).Where(sheet => calculation.ClearCachedResults || calculation.RecalculateOnOpen
            || encoding.Cells.Keys.Concat(encoding.FormulaCaches.Keys).Any(key => key.Sheet == (structure == null ? sheet.Name : structure.CurrentName(sheet.Name))))) {
            token.ThrowIfCancellationRequested();
            string currentName = structure?.CurrentName(sheet.Name) ?? sheet.Name;
            int begin = _records.ToList().FindIndex(record => record.Offset == sheet.StreamOffset);
            int dimension = -1, index = -1, depth = 0;
            var cells = new SortedDictionary<(ushort Row, ushort Column), List<BiffRecord>>();
            var rows = new SortedDictionary<ushort, byte[]>();
            List<BiffRecord>? previousCell = null;
            for (int i = begin; i < _records.Count; i++) {
                var record = _records[i]; byte[] payload = record.Payload;
                if (record.Type == 0x0809) { depth++; continue; }
                if (record.Type == 0x000a) { if (--depth == 0) break; continue; }
                if (depth != 1) continue;
                if (record.Type == 0x0200) dimension = i;
                if (record.Type == 0x020b) index = i;
                if (record.Type == 0x0208) { rows.Add(BiffRecordReader.ReadUInt16(payload, 0), payload); layout.Replace[i] = Array.Empty<OutputRecord>(); }
                else if (record.Type == 0x00d7) layout.Replace[i] = Array.Empty<OutputRecord>();
                else if (record.Type is 0x00bd or 0x00be) {
                    int stride = record.Type == 0x00bd ? 6 : 2;
                    int count = (payload.Length - 6) / stride;
                    for (int cell = 0; cell < count; cell++) {
                        var data = new byte[record.Type == 0x00bd ? 10 : 6];
                        Buffer.BlockCopy(payload, 0, data, 0, 2);
                        int column = BiffRecordReader.ReadUInt16(payload, 2) + cell;
                        data[2] = (byte)column; data[3] = (byte)(column >> 8);
                        Buffer.BlockCopy(payload, 4 + cell * stride, data, 4, stride);
                        Add(new BiffRecord(record.Type == 0x00bd ? (ushort)0x027e : (ushort)0x0201, 0, data));
                    }
                    layout.Replace[i] = Array.Empty<OutputRecord>();
                } else if (record.Type is 0x0201 or 0x0203 or 0x027e or 0x00fd or 0x0205 or 0x0006 or 0x0204 or 0x00d6) {
                    Add(record); layout.Replace[i] = Array.Empty<OutputRecord>();
                } else if (previousCell != null && record.Type is 0x0207 or 0x04bc or 0x0221 or 0x0236 or 0x003c) {
                    previousCell.Add(record); layout.Replace[i] = Array.Empty<OutputRecord>();
                } else if (record.Type != 0x0208 && record.Type != 0x00d7) previousCell = null;
            }
            if (dimension < 0 || index < 0) throw new NotSupportedException("The worksheet requires native Dimensions and Index records.");
            foreach (var edit in encoding.Cells.Where(item => item.Key.Sheet == currentName)) {
                var address = A1.ParseCellRef(edit.Key.Address);
                var key = (checked((ushort)(address.Row - 1)), checked((ushort)(address.Col - 1)));
                if (edit.Value.Count == 0) cells.Remove(key);
                else cells[key] = edit.Value.ToList();
            }
            ApplyFormulaCaches(document, currentName, cells, encoding, calculation, structure, token);
            foreach (var row in cells.Keys.Select(key => key.Row).Distinct()) if (!rows.ContainsKey(row)) rows.Add(row, Array.Empty<byte>());
            var rowNumbers = rows.Keys.ToArray();
            uint first = rowNumbers.Length == 0 ? 0U : rowNumbers[0];
            uint after = rowNumbers.Length == 0 ? 0 : (uint)rowNumbers[rowNumbers.Length - 1] + 1;
            int blockCount = checked((int)((after - first + 31) / 32));
            byte[] dimensions = new byte[14]; WriteUInt32(dimensions, 0, first); WriteUInt32(dimensions, 4, after);
            ushort firstColumn = cells.Count == 0 ? (ushort)0 : cells.Keys.Min(key => key.Column);
            ushort afterColumn = cells.Count == 0 ? (ushort)0 : checked((ushort)(cells.Keys.Max(key => key.Column) + 1));
            dimensions[8] = (byte)firstColumn; dimensions[9] = (byte)(firstColumn >> 8);
            dimensions[10] = (byte)afterColumn; dimensions[11] = (byte)(afterColumn >> 8);
            layout.Replace[dimension] = new[] { new OutputRecord(dimension, 0x0200, dimensions) };
            byte[] indexPayload = new byte[16 + 4 * blockCount]; WriteUInt32(indexPayload, 4, first); WriteUInt32(indexPayload, 8, after);
            if (_records[index].Payload.Length >= 16) Buffer.BlockCopy(_records[index].Payload, 12, indexPayload, 12, 4);
            layout.Replace[index] = new[] { new OutputRecord(index, 0x020b, indexPayload) };
            var dbKeys = new List<int>(); var table = new List<OutputRecord>(); long position = 0;
            void Emit(OutputRecord record) { table.Add(record); position += 4 + record.Payload.Length; }
            for (uint block = first; block < after; block += 32) {
                var blockRows = rowNumbers.Where(row => row >= block && row < block + 32).ToArray();
                long firstRow = position;
                foreach (ushort row in blockRows) {
                    ushort[] columns = cells.Keys.Where(key => key.Row == row).Select(key => key.Column).ToArray();
                    Emit(layout.New(0x0208, LegacyXlsWriter.EncodePreservedRow(row, columns, rows[row].Length == 0 ? null : rows[row])));
                }
                var firstCells = new List<long>();
                foreach (ushort row in blockRows) {
                    var rowCells = cells.Where(item => item.Key.Row == row).ToArray();
                    if (rowCells.Length > 0) firstCells.Add(position);
                    foreach (var cell in rowCells) foreach (var record in cell.Value) Emit(layout.New(record.Type, record.Payload));
                }
                var db = layout.New(0x00d7, LegacyXlsWriter.EncodePreservedDbCell(firstRow, firstCells, position));
                dbKeys.Add(db.Key); Emit(db);
            }
            foreach (var record in table) layout.InsertBefore(dimension + 1, record);
            layout.IndexDbCells.Add(index, dbKeys);

            void Add(BiffRecord record) {
                var key = (BiffRecordReader.ReadUInt16(record.Payload, 0), BiffRecordReader.ReadUInt16(record.Payload, 2));
                if (cells.ContainsKey(key)) throw new NotSupportedException("Duplicate native cells cannot be edited safely.");
                cells.Add(key, previousCell = new List<BiffRecord> { record });
            }
        }
        return layout;
    }

    private sealed class EditingLayout {
        private int _key = int.MinValue;
        internal Dictionary<int, IReadOnlyList<OutputRecord>> Replace { get; } = new();
        internal Dictionary<int, List<OutputRecord>> Before { get; } = new();
        internal Dictionary<(int Record, int Offset), int> PointerTargets { get; } = new();
        internal Dictionary<int, List<int>> IndexDbCells { get; } = new();
        internal OutputRecord New(ushort type, byte[] payload) => new(_key++, type, payload);
        internal void InsertBefore(int index, OutputRecord record) {
            if (!Before.TryGetValue(index, out var list)) Before.Add(index, list = new());
            list.Add(record);
        }
    }
}
