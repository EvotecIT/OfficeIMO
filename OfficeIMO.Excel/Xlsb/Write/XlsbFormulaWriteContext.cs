using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Appends self-reference slots for edited formulas without changing the indexes used by retained native formulas.</summary>
internal sealed class XlsbFormulaWriteContext {
    private readonly ExcelDocument _document;
    private readonly XlsbWorkbook _source;
    private readonly XlsbSheetStructure _structure;
    private LegacyXlsExternSheetTable? _links;
    private LegacyXlsFormulaNameIndex? _names;
    private readonly Dictionary<ushort, ushort> _mapped = new();
    private readonly List<(int First, int Last)> _added = new();
    private int _selfLink = -1;

    internal XlsbFormulaWriteContext(ExcelDocument document, XlsbWorkbook source, XlsbSheetStructure structure) {
        _document = document; _source = source; _structure = structure;
    }

    internal bool TryEncode(string text, string sheetName, out byte[] payload, out string? reason) {
        var sheets = _document.Sheets.ToArray();
        _links ??= LegacyXlsExternSheetTable.Create(_document, sheets);
        if (_links.SupportingLinkRecords.Count > 1) throw new NotSupportedException("New external-workbook XLSB formulas require an external-link writer.");
        if (_names == null) {
            var global = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            var local = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < _source.DefinedNames.Count; index++) {
                var name = _source.DefinedNames[index]; uint? mapped = _structure.NameIndexes[checked((uint)index + 1)];
                if (!mapped.HasValue) continue; uint id = mapped.Value;
                if (name.LocalSheetIndex == uint.MaxValue) global[name.Name] = id;
                else {
                    var binding = _structure.Bindings.FirstOrDefault(item => item.Original?.NativeIndex == name.LocalSheetIndex);
                    int current = binding == null ? -1 : Array.FindIndex(sheets, sheet => sheet.Name == binding.Name);
                    if (current >= 0) local[LegacyXlsFormulaNameIndex.CreateLocalKey(current, name.Name)] = id;
                }
            }
            _names = new LegacyXlsFormulaNameIndex(global, local, _links);
        }
        int sheetIndex = Array.FindIndex(sheets, sheet => sheet.Name == sheetName);
        bool result = XlsbFormulaEncoder.TryEncode(text, _names, sheetIndex,
            MapReference, out payload, out reason);
        return result;
    }

    internal XlsbWriteCell RebindSourceCell(XlsbWriteCell cell) {
        if (cell.SourceRecordType == null || cell.SourceRecordData == null || !_structure.NameIndexes.Any(item => item.Value != item.Key)) return cell;
        byte[] data = XlsbFormulaRecordRebinder.Rebind(cell.SourceRecordType.Value, cell.SourceRecordData, _structure.NameIndexes);
        return new XlsbWriteCell(cell.Row, cell.Column, cell.StyleIndex, cell.Kind, cell.Value, cell.FormulaPayload, cell.SourceRecordType, data);
    }

    internal byte[] RebindSourceFormula(byte[] payload) => _structure.NameIndexes.Any(item => item.Value != item.Key)
        ? XlsbFormulaRecordRebinder.RebindCellFormulaPayload(payload, _structure.NameIndexes) : payload;

    private ushort MapReference(ushort generated) {
        if (_mapped.TryGetValue(generated, out ushort mapped)) return mapped;
        if (_selfLink < 0) {
            _selfLink = Enumerable.Range(0, _source.SelfSupportingLinks.Count).Where(index => _source.SelfSupportingLinks[index]).DefaultIfEmpty(-1).First();
            if (_selfLink < 0) _selfLink = _source.SelfSupportingLinks.Count;
        }
        var sheets = _document.Sheets.ToArray(); int from = 2 + generated * 6;
        if (from > _links!.Payload.Length - 6) throw new InvalidDataException("Invalid generated XLSB reference index.");
        int first = _links.Payload[from + 2] | _links.Payload[from + 3] << 8;
        int last = _links.Payload[from + 4] | _links.Payload[from + 5] << 8;
        if (first >= sheets.Length || last >= sheets.Length) throw new InvalidDataException("Invalid generated self-reference.");
        first = _structure.Bindings.Single(binding => binding.Name == sheets[first].Name).Index;
        last = _structure.Bindings.Single(binding => binding.Name == sheets[last].Name).Index;
        int Rebind(int index) => _structure.Bindings.FirstOrDefault(binding => binding.Original?.NativeIndex == index)?.Index ?? -1;
        for (int index = 0; index < _source.ExternalSheetReferences.Count; index++) {
            var old = _source.ExternalSheetReferences[index];
            if (old.SupportingLinkIndex == _selfLink && Rebind(old.FirstSheetIndex) == first && Rebind(old.LastSheetIndex) == last) {
                mapped = checked((ushort)index); _mapped.Add(generated, mapped); return mapped;
            }
        }
        int added = _added.FindIndex(item => item.First == first && item.Last == last);
        if (added < 0) { added = _added.Count; _added.Add((first, last)); }
        mapped = checked((ushort)(_source.ExternalSheetReferences.Count + added)); _mapped.Add(generated, mapped); return mapped;
    }

    internal byte[] CompleteWorkbook(byte[] original, long maximumBytes) {
        if (_added.Count == 0) return original;
        using var input = new MemoryStream(original); var records = XlsbRecordReader.ReadAll(input);
        var tables = records.Where(record => record.Type == 362).ToArray();
        if (tables.Length > 1) throw new NotSupportedException("The native XLSB external-sheet table must be consolidated before appending references.");
        byte[] old = tables.Length == 0 ? new byte[4] : tables[0].Data;
        int oldCount = checked((int)Read(old, 0));
        if (old.Length != 4 + checked(oldCount * 12) || oldCount != _source.ExternalSheetReferences.Count || oldCount + _added.Count > ushort.MaxValue)
            throw new InvalidDataException("The native XLSB external-sheet table is invalid or full.");
        byte[] combined = new byte[checked(old.Length + _added.Count * 12)]; Buffer.BlockCopy(old, 0, combined, 0, old.Length);
        Put(combined, 0, checked((uint)(oldCount + _added.Count)));
        for (int index = 0; index < _added.Count; index++) {
            int to = old.Length + index * 12;
            Put(combined, to, checked((uint)_selfLink));
            Put(combined, to + 4, checked((uint)_added[index].First)); Put(combined, to + 8, checked((uint)_added[index].Last));
        }
        bool hasExternals = records.Any(record => record.Type == 353), written = false;
        using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
        void WriteLinks() {
            if (_selfLink == _source.SelfSupportingLinks.Count) XlsbRecordWriter.Write(output, 357);
            XlsbRecordWriter.Write(output, 362, combined); written = true;
        }
        foreach (var record in records) {
            if (record.Type == 362) { WriteLinks(); continue; }
            if (record.Type == 354 && !written) WriteLinks();
            XlsbRecordWriter.Write(output, record.Type, record.Data);
            if (record.Type == 144 && !hasExternals) {
                XlsbRecordWriter.Write(output, 353); WriteLinks(); XlsbRecordWriter.Write(output, 354);
            }
        }
        if (!written) throw new InvalidDataException("The native XLSB workbook has no insertion point for external references.");
        return output.ToArray();
    }

    private static uint Read(byte[] bytes, int offset) {
        if (offset > bytes.Length - 4) throw new InvalidDataException("A native XLSB external-sheet count is truncated.");
        return (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
    }
    private static void Put(byte[] bytes, int offset, uint value) { for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (index * 8)); }
}
