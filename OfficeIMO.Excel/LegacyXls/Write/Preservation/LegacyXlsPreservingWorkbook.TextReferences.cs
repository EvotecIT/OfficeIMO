using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private void RewriteTextSheetReferences(EditingLayout layout, LegacyXlsSheetStructure structure, CancellationToken token) {
        var names = structure.Bindings.Where(binding => binding.Original != null && binding.Name != binding.Original.Name)
            .ToDictionary(binding => binding.Original!.Name, binding => binding.Name, StringComparer.OrdinalIgnoreCase);
        if (names.Count == 0) return;
        for (int index = 0; index < _records.Count; index++) {
            token.ThrowIfCancellationRequested(); var record = _records[index];
            if (layout.Replace.TryGetValue(index, out var replaced) && replaced.Count == 0) continue;
            byte[] payload = record.Payload, updated = payload;
            if (record.Type == 0x01b8) updated = BiffHyperlinkReader.RenameInternalSheet(payload, names);
            else if (record.Type == 0x0051 && BiffDataConsolidationReferenceReader.TryRead(record, new(), out var reference)
                && reference?.SourceKind == LegacyXlsDataConsolidationSourceKind.SelfReference && names.TryGetValue(reference.Source, out string? current)) {
                int offset = 8; BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, BiffRecordReader.ReadUInt16(payload, 6));
                string source = "\u0002" + current; byte[] text = Encoding.Unicode.GetBytes(source);
                updated = new byte[checked(9 + text.Length + payload.Length - offset)];
                Buffer.BlockCopy(payload, 0, updated, 0, 6); LegacyXlsSheetStructure.Put16(updated, 6, source.Length); updated[8] = 1;
                Buffer.BlockCopy(text, 0, updated, 9, text.Length); Buffer.BlockCopy(payload, offset, updated, 9 + text.Length, payload.Length - offset);
            }
            if (!ReferenceEquals(updated, payload)) layout.Replace[index] = new[] { new OutputRecord(index, record.Type, updated) };
        }
    }
}
