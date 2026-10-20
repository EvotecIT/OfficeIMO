using OfficeIMO.Excel.LegacyXls.Biff;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private void ApplyCalculationSettings(ExcelDocument document, EditingLayout layout, LegacyXlsSheetStructure? structure, CancellationToken token) {
        var settings = LegacyXlsWriter.EncodeCalculationSettings(document);
        int globalEnd = _records.ToList().FindIndex(record => record.Type == 0x000a);
        foreach (var record in settings.Globals) {
            var positions = Enumerable.Range(0, globalEnd).Where(index => _records[index].Type == record.Type).ToArray();
            foreach (int index in positions) layout.Replace[index] = new[] { new OutputRecord(index, record.Type, record.Payload) };
            if (positions.Length == 0) layout.InsertBefore(globalEnd, layout.New(record.Type, record.Payload));
        }
        foreach (var sheet in _model.Worksheets.Where(sheet => structure == null || structure.CurrentName(sheet.Name) != null)) {
            token.ThrowIfCancellationRequested();
            int begin = _records.ToList().FindIndex(record => record.Offset == sheet.StreamOffset), depth = 0, insertion = begin + 1;
            var positions = new Dictionary<ushort, List<int>>();
            for (int index = begin; index < _records.Count; index++) {
                var record = _records[index];
                if (record.Type == 0x0809) { depth++; continue; }
                if (record.Type == 0x000a) { if (--depth == 0) break; continue; }
                if (depth != 1) continue;
                if (record.Type == 0x020b) insertion = index + 1;
                if (!positions.TryGetValue(record.Type, out var found)) positions.Add(record.Type, found = new());
                found.Add(index);
            }
            foreach (var record in settings.Worksheets) {
                if (positions.TryGetValue(record.Type, out var found))
                    foreach (int index in found) layout.Replace[index] = new[] { new OutputRecord(index, record.Type, record.Payload) };
                else layout.InsertBefore(insertion, layout.New(record.Type, record.Payload));
            }
        }
    }

    private static void ApplyFormulaCaches(ExcelDocument document, string sheetName,
        SortedDictionary<(ushort Row, ushort Column), List<BiffRecord>> cells, LegacyXlsEditEncoding encoding,
        ExcelNativeCalculationPolicy policy, LegacyXlsSheetStructure? structure, CancellationToken token) {
        var sheet = document.GetSheet(sheetName);
        foreach (var cell in cells.Keys.ToArray()) {
            token.ThrowIfCancellationRequested(); var records = cells[cell];
            if (records.Count == 0 || records[0].Type != 0x0006) continue;
            string address = A1.CellReference(cell.Row + 1, cell.Column + 1);
            if (encoding.FormulaCaches.TryGetValue((sheetName, address), out var current))
                records = cells[cell] = LegacyXlsWriter.EncodePreservedFormulaCache(sheet, current, records, document.DateSystem, structure?.NameIndexes).ToList();
            if (!policy.ClearCachedResults && !policy.RecalculateOnOpen) continue;
            byte[] payload = (byte[])records[0].Payload.Clone();
            if (payload.Length < 22) throw new InvalidDataException("The native XLS formula cache is truncated.");
            if (policy.ClearCachedResults) {
                Array.Clear(payload, 6, 8); payload[6] = 3; payload[12] = payload[13] = 0xFF;
                bool stringContinuation = false; var kept = new List<BiffRecord>();
                foreach (var record in records.Skip(1)) {
                    if (record.Type == 0x0207) { stringContinuation = true; continue; }
                    if (stringContinuation && record.Type == 0x003c) continue;
                    stringContinuation = false; kept.Add(record);
                }
                records.RemoveRange(1, records.Count - 1); records.AddRange(kept);
            }
            if (policy.RecalculateOnOpen) payload[14] |= 1; // Formula.fAlwaysCalc, reserved bits remain unchanged.
            records[0] = new BiffRecord(0x0006, records[0].Offset, payload);
        }
    }
}
