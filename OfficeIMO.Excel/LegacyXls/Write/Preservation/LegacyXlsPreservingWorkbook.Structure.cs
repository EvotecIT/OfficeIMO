using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.Utilities;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private readonly Dictionary<string, string> _allocatedCodeNames = new(StringComparer.Ordinal);

    private OfficeCompoundFile PrepareVbaHosts(LegacyXlsSheetStructure? structure, long maximumBytes, CancellationToken token) {
        if (structure == null) return _compound;
        if (_model.UnsupportedSheets.Count != 0)
            throw new NotSupportedException("Native worksheet structure editing requires all original sheet types to be projected.");
        if (!_compound.Streams.Keys.Any(path => path.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase))) return _compound;
        token.ThrowIfCancellationRequested();
        var storage = OfficeCompoundStorage.Extract(_compound, "_VBA_PROJECT_CUR");
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(storage, new Dictionary<string, byte[]>(), maxOutputBytes: maximumBytes, cancellationToken: token);
        byte[] edited = ExcelLegacyVbaHostEditor.Reconcile(bytes, structure.Bindings,
            _model.Worksheets.ToDictionary(sheet => sheet.Name, sheet => sheet.CodeName, StringComparer.Ordinal), _allocatedCodeNames, structure.NewCodeNames);
        if (ReferenceEquals(bytes, edited)) return _compound;
        if (!OfficeCompoundFileReader.TryRead(edited, out OfficeCompoundFile? replacement, out string? error) || replacement == null)
            throw new InvalidDataException("The edited VBA storage is invalid: " + error);
        return OfficeCompoundStorage.Replace(_compound, "_VBA_PROJECT_CUR", replacement);
    }

    private void ApplyStructure(EditingLayout layout, LegacyXlsSheetStructure structure, LegacyXlsEditEncoding encoding, CancellationToken token) {
        foreach (var change in structure.GlobalChanges) {
            if (_records[change.Key].Type == 0x0017) continue; // The encoding plan appends new references to the remapped original table.
            layout.Replace[change.Key] = change.Value == null ? Array.Empty<OutputRecord>()
                : new[] { new OutputRecord(change.Key, _records[change.Key].Type, change.Value) };
        }
        var bounds = _records.Select((record, index) => (record, index)).Where(item => item.record.Type == 0x0085).ToArray();
        if (bounds.Length == 0) throw new NotSupportedException("The native workbook has no sheet bindings.");
        var sourceBounds = _nativeSheets.ToDictionary(item => item.Key, item => bounds[item.Value], StringComparer.Ordinal);
        var sourceBofs = _records.Select((record, index) => (record, index)).Where(item => item.record.Type == 0x0809)
            .ToDictionary(item => item.record.Offset, item => item.index);
        foreach (var item in bounds) layout.Replace[item.index] = Array.Empty<OutputRecord>();
        int appendAt = _records.Count;
        while (appendAt > 0 && _records[appendAt - 1].Type == 0 && _records[appendAt - 1].Payload.Length == 0) appendAt--;
        foreach (var binding in structure.Bindings) {
            token.ThrowIfCancellationRequested();
            int headerKey, bofKey; byte sheetType;
            if (binding.Original != null) {
                var native = sourceBounds[binding.Original.Name];
                headerKey = native.index; sheetType = native.record.Payload[5];
                bofKey = sourceBofs[checked((int)BiffRecordReader.ReadUInt32(native.record.Payload, 0))];
            } else {
                var records = encoding.NewWorksheets[binding.Name];
                var output = records.Select(record => layout.New(record.Type, record.Payload)).ToArray();
                var offsets = records.Select((record, index) => (record, index)).Where(item => item.record.Offset >= 0)
                    .ToDictionary(item => item.record.Offset, item => output[item.index].Key);
                for (int i = 0; i < output.Length; i++) if (output[i].Type == 0x020b) {
                    for (int field = 12; field < output[i].Payload.Length; field += 4) {
                        int pointer = checked((int)BiffRecordReader.ReadUInt32(output[i].Payload, field));
                        if (pointer != 0) layout.PointerTargets[(output[i].Key, field)] = offsets[pointer];
                    }
                }
                foreach (var record in output) layout.InsertBefore(appendAt, record);
                headerKey = layout.New(0x0085, Array.Empty<byte>()).Key; bofKey = output[0].Key; sheetType = 0;
            }
            byte[] name = Encoding.Unicode.GetBytes(binding.Name);
            byte[] payload = new byte[8 + name.Length];
            payload[4] = binding.Element.State?.Value == SheetStateValues.VeryHidden ? (byte)2
                : binding.Element.State?.Value == SheetStateValues.Hidden ? (byte)1 : (byte)0;
            payload[5] = sheetType; payload[6] = checked((byte)binding.Name.Length); payload[7] = 1;
            Buffer.BlockCopy(name, 0, payload, 8, name.Length);
            layout.InsertBefore(bounds[0].index, new OutputRecord(headerKey, 0x0085, payload));
            layout.PointerTargets[(headerKey, 0)] = bofKey;
        }
        var retained = new HashSet<string>(structure.Bindings.Where(binding => binding.Original != null).Select(binding => binding.Original!.Name), StringComparer.Ordinal);
        foreach (var sheet in _model.Worksheets.Where(sheet => !retained.Contains(sheet.Name))) {
            int start = sourceBofs[sheet.StreamOffset]; int depth = 0;
            for (int i = start; i < _records.Count; i++) {
                layout.Replace[i] = Array.Empty<OutputRecord>();
                if (_records[i].Type == 0x0809) depth++;
                if (_records[i].Type == 0x000a && --depth == 0) break;
            }
        }
        if (structure.NameIndexes.Any(item => item.Value != item.Key)) RewriteNameReferences(layout, structure.NameIndexes, token);
    }
}
