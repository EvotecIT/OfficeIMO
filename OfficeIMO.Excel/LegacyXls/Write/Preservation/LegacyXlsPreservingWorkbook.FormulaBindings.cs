using OfficeIMO.Excel.LegacyXls.Biff;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private void RewriteNameReferences(EditingLayout layout, IReadOnlyDictionary<uint, uint?> names, CancellationToken token) {
        var nativePayloads = new HashSet<byte[]>(_records.Select(record => record.Payload));
        OutputRecord Rebind(OutputRecord record) {
            if (record.Key < 0 && !nativePayloads.Contains(record.Payload)) return record;
            var ranges = GetFormulaRanges(record.Type, record.Payload);
            if (ranges.Count == 0) return record;
            byte[] payload = (byte[])record.Payload.Clone();
            foreach (var range in ranges) BiffFormulaTokenScanner.RewriteNames(payload, range.Offset, range.Length, names);
            return record with { Payload = payload };
        }
        for (int index = 0; index < _records.Count; index++) {
            token.ThrowIfCancellationRequested();
            if (layout.Replace.TryGetValue(index, out var replacement)) layout.Replace[index] = replacement.Select(Rebind).ToArray();
            else {
                var original = new OutputRecord(index, _records[index].Type, _records[index].Payload);
                var rebound = Rebind(original);
                if (!ReferenceEquals(original.Payload, rebound.Payload)) layout.Replace[index] = new[] { rebound };
            }
        }
        foreach (var list in layout.Before.Values) for (int index = 0; index < list.Count; index++) list[index] = Rebind(list[index]);
    }

    private static IReadOnlyList<(int Offset, int Length)> GetFormulaRanges(ushort type, byte[] payload) {
        var ranges = new List<(int, int)>();
        void Prefix(int position, int reserved = 0) {
            if (position > payload.Length - 2) throw new InvalidDataException("A native formula length field is truncated.");
            int length = BiffRecordReader.ReadUInt16(payload, position);
            ranges.Add((position + 2 + reserved, length));
        }
        switch (type) {
            case 0x0006: Prefix(20); break;
            case 0x04bc: Prefix(8); break;
            case 0x0221: Prefix(12); break;
            case 0x1051: Prefix(6); break;
            case 0x00f8: Prefix(0); break;
            case 0x0018: {
                if (payload.Length < 14) throw new InvalidDataException("A native defined name is truncated.");
                int offset = 14; BiffStringReader.ReadUnicodeStringNoCch(payload, ref offset, payload[3]);
                ranges.Add((offset, BiffRecordReader.ReadUInt16(payload, 4))); break;
            }
            case 0x01b1: {
                if (payload.Length < 6) throw new InvalidDataException("A native conditional rule is truncated.");
                int first = BiffRecordReader.ReadUInt16(payload, 2), second = BiffRecordReader.ReadUInt16(payload, 4);
                ranges.Add((payload.Length - first - second, first)); ranges.Add((payload.Length - second, second)); break;
            }
            case 0x01be: {
                int offset = 4;
                for (int index = 0; index < 4; index++) BiffStringReader.ReadUnicodeString(payload, ref offset);
                Prefix(offset, 2); offset += 4 + BiffRecordReader.ReadUInt16(payload, offset);
                Prefix(offset, 2); break;
            }
        }
        return ranges;
    }
}
