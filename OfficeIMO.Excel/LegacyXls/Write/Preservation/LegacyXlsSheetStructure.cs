using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.LegacyXls.Write;

/// <summary>Maintains stable native links while the physical sheet collection changes.</summary>
internal sealed class LegacyXlsSheetStructure {
    internal LegacyXlsSheetStructure(IReadOnlyList<ExcelLegacySheetBindings.Binding> bindings, IReadOnlyDictionary<string, int> originalIndexes) {
        Bindings = bindings;
        if (bindings.Count == 0 || bindings.Count > 32767) throw new NotSupportedException("Native XLS requires 1–32767 sheets.");
        if (bindings.Select(binding => binding.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != bindings.Count)
            throw new NotSupportedException("Native worksheet names must be unique.");
        foreach (var binding in bindings)
            if (binding.Name.Length == 0 || binding.Name.Length > 31) throw new NotSupportedException("Native worksheet names require 1–31 characters.");
        CurrentIndexes = bindings.ToDictionary(binding => binding.Name, binding => binding.Index, StringComparer.Ordinal);
        OriginalIndexes = originalIndexes;
    }
    internal IReadOnlyList<ExcelLegacySheetBindings.Binding> Bindings { get; }
    internal IReadOnlyDictionary<string, int> CurrentIndexes { get; }
    internal IReadOnlyDictionary<string, int> OriginalIndexes { get; }
    internal Dictionary<uint, uint?> NameIndexes { get; } = new();
    internal Dictionary<int, byte[]?> GlobalChanges { get; } = new();
    internal Dictionary<string, string> NewCodeNames { get; } = new(StringComparer.Ordinal);
    internal int Map(int index) => Bindings.FirstOrDefault(binding => binding.Original?.NativeIndex == index)?.Index ?? -1;
    internal string? CurrentName(string original) => Bindings.FirstOrDefault(binding => binding.Original?.Name == original)?.Name;

    internal IReadOnlyList<BiffRecord> PrepareGlobals(IReadOnlyList<BiffRecord> records) {
        var selfBooks = new HashSet<int>();
        int supBook = 0; uint nameIndex = 0, keptNames = 0;
        for (int i = 0; i < records.Count && records[i].Type != 0x000a; i++) {
            var record = records[i]; byte[] payload = record.Payload;
            if (record.Type == 0x01ae) {
                if (payload.Length == 4 && payload[2] == 1 && payload[3] == 4) {
                    selfBooks.Add(supBook); payload = (byte[])payload.Clone(); Put16(payload, 0, Bindings.Count); GlobalChanges[i] = payload;
                }
                supBook++;
            } else if (record.Type == 0x0018) {
                if (payload.Length < 14) throw new NotSupportedException("A native defined name is truncated.");
                nameIndex++; int scope = BiffRecordReader.ReadUInt16(payload, 8);
                int mapped = scope == 0 ? -1 : Map(scope - 1);
                if (scope > 0 && mapped < 0) { NameIndexes[nameIndex] = null; GlobalChanges[i] = null; }
                else { NameIndexes[nameIndex] = ++keptNames; if (scope > 0) { payload = (byte[])payload.Clone(); Put16(payload, 8, mapped + 1); GlobalChanges[i] = payload; } }
            }
        }
        for (int i = 0; i < records.Count && records[i].Type != 0x000a; i++) {
            var record = records[i]; byte[] payload = record.Payload;
            if (record.Type == 0x0017) {
                if (payload.Length < 2 || payload.Length != 2 + BiffRecordReader.ReadUInt16(payload, 0) * 6)
                    throw new NotSupportedException("A continued native sheet-reference table requires consolidation.");
                payload = (byte[])payload.Clone();
                for (int offset = 2; offset < payload.Length; offset += 6) {
                    if (!selfBooks.Contains(BiffRecordReader.ReadUInt16(payload, offset))) continue;
                    short first = unchecked((short)BiffRecordReader.ReadUInt16(payload, offset + 2));
                    short last = unchecked((short)BiffRecordReader.ReadUInt16(payload, offset + 4));
                    if (first < 0 || last < 0) continue;
                    var surviving = Bindings.Where(binding => binding.Original != null
                        && binding.Original.NativeIndex >= Math.Min(first, last) && binding.Original.NativeIndex <= Math.Max(first, last))
                        .OrderBy(binding => binding.Original!.NativeIndex).ToArray();
                    int newFirst = surviving.Length == 0 ? 65535 : surviving[0].Index;
                    int newLast = surviving.Length == 0 ? 65535 : surviving[surviving.Length - 1].Index;
                    Put16(payload, offset + 2, Math.Min(newFirst, newLast)); Put16(payload, offset + 4, Math.Max(newFirst, newLast));
                }
                GlobalChanges[i] = payload;
            } else if (record.Type == 0x013d) {
                if (payload.Length != OriginalIndexes.Count * 2) throw new NotSupportedException("The native tab-identity table is incomplete.");
                var used = new HashSet<int>(Enumerable.Range(0, payload.Length / 2).Select(index => (int)BiffRecordReader.ReadUInt16(payload, index * 2)));
                var rewritten = new byte[Bindings.Count * 2]; int next = 1;
                foreach (var binding in Bindings) {
                    int tab;
                    if (binding.Original != null) tab = BiffRecordReader.ReadUInt16(payload, binding.Original.NativeIndex * 2);
                    else { while (used.Contains(next)) next++; tab = next; used.Add(next++); }
                    Put16(rewritten, binding.Index * 2, tab);
                }
                GlobalChanges[i] = rewritten;
            } else if (record.Type == 0x003d && payload.Length >= 18) {
                payload = (byte[])payload.Clone();
                Put16(payload, 10, Math.Max(0, Map(BiffRecordReader.ReadUInt16(payload, 10))));
                Put16(payload, 12, Math.Max(0, Map(BiffRecordReader.ReadUInt16(payload, 12))));
                Put16(payload, 14, Math.Max(1, Math.Min(Bindings.Count, BiffRecordReader.ReadUInt16(payload, 14))));
                GlobalChanges[i] = payload;
            }
        }
        var result = new List<BiffRecord>(records.Count);
        for (int index = 0; index < records.Count; index++) {
            var record = records[index];
            if (GlobalChanges.TryGetValue(index, out var changed)) {
                if (changed != null) result.Add(new BiffRecord(record.Type, record.Offset, changed));
            } else result.Add(record);
        }
        return result;
    }

    internal static void Put16(byte[] payload, int offset, int value) {
        ushort encoded = checked((ushort)value); payload[offset] = (byte)encoded; payload[offset + 1] = (byte)(encoded >> 8);
    }
}
