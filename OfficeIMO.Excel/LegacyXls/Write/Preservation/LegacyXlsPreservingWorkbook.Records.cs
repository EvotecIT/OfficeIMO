using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private byte[] RewriteRecords(Dictionary<int, byte[]> replacements, IReadOnlyList<string> strings,
        byte[]? filePass, long maximumBytes, CancellationToken cancellationToken) {
        var output = new List<OutputRecord>(_records.Count);
        var locations = _stringLocations.Select(position => (RecordKey: _sstIndex + position.SegmentIndex, position.Offset)).ToList();
        int syntheticKey = -1;
        IReadOnlyList<byte[]> appended = Array.Empty<byte[]>();
        int extKey = _extSstIndex;
        uint stringCount = checked(_stringCount + (uint)strings.Count);
        if (strings.Count != 0) {
            var encoded = LegacyXlsWriter.EncodePreservedStringAppend(strings, maximumBytes, cancellationToken);
            appended = encoded.Payloads;
            foreach (var position in encoded.Locations) locations.Add((-1 - position.SegmentIndex, position.Offset));
            syntheticKey -= appended.Count;
            byte[] sst = (byte[])_records[_sstIndex].Payload.Clone(); WriteUInt32(sst, 4, stringCount); replacements[_sstIndex] = sst;
        }
        if (_sstIndex >= 0 && extKey < 0) extKey = syntheticKey--;
        int bucketSize = Math.Max(checked((int)(stringCount / 128)) + 1, 8);
        int bucketCount = checked((int)((stringCount + bucketSize - 1) / bucketSize));
        byte[]? newExtSst = _sstIndex < 0 ? null : new byte[checked(2 + 8 * bucketCount)];
        ushort encodedBucketSize = checked((ushort)bucketSize);
        if (newExtSst != null) { newExtSst[0] = (byte)encodedBucketSize; newExtSst[1] = (byte)(encodedBucketSize >> 8); }
        for (int i = 0; i < _records.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (_sstIndex >= 0 && i == _sstEnd) {
                for (int append = 0; append < appended.Count; append++) output.Add(new OutputRecord(-1 - append, 0x003c, appended[append]));
                if (_extSstIndex < 0) output.Add(new OutputRecord(extKey, 0x00ff, newExtSst!));
            }
            var record = _records[i];
            byte[] payload = newExtSst != null && i == _extSstIndex ? newExtSst : replacements.TryGetValue(i, out var replacement) ? replacement : record.Payload;
            output.Add(new OutputRecord(i, record.Type, payload));
            if (i == 0 && filePass != null) output.Add(new OutputRecord(syntheticKey--, 0x002f, filePass));
        }
        var newOffsets = new Dictionary<int, int>(); var oldToNew = new Dictionary<int, int>();
        long length = 0;
        foreach (var record in output) {
            if (length > maximumBytes || length > int.MaxValue) throw OfficeOutputLimit.Create("Native XLS preservation exceeds the output materialization limit.");
            newOffsets.Add(record.Key, checked((int)length));
            if (record.Key >= 0) oldToNew.Add(_records[record.Key].Offset, checked((int)length));
            length = checked(length + 4 + record.Payload.Length);
        }
        if (length > maximumBytes || length > int.MaxValue) throw OfficeOutputLimit.Create("Native XLS preservation exceeds the output materialization limit.");
        uint Relocate(uint source) => source == 0 ? 0 : oldToNew.TryGetValue(checked((int)source), out int target) ? (uint)target
            : throw new NotSupportedException("The source workbook contains an unresolved native record pointer.");
        using var stream = new OfficeBoundedMemoryStream(maximumBytes);
        foreach (var record in output) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] payload = record.Payload;
            if (record.Type == 0x0085) { payload = (byte[])payload.Clone(); WriteUInt32(payload, 0, Relocate(BiffRecordReader.ReadUInt32(payload, 0))); }
            else if (record.Type == 0x020b) {
                if (payload.Length < 16 || (payload.Length - 16) % 4 != 0) throw new NotSupportedException("Invalid native worksheet Index record.");
                payload = (byte[])payload.Clone();
                for (int offset = 12; offset < payload.Length; offset += 4) WriteUInt32(payload, offset, Relocate(BiffRecordReader.ReadUInt32(payload, offset)));
            } else if (record.Type == 0x00ff) {
                payload = (byte[])payload.Clone();
                if (newExtSst != null && record.Key == extKey) {
                    for (int bucket = 0; bucket < bucketCount; bucket++) {
                        var location = locations[bucket * bucketSize]; int offset = 2 + bucket * 8;
                        WriteUInt32(payload, offset, checked((uint)(newOffsets[location.RecordKey] + location.Offset + 4)));
                        ushort relative = checked((ushort)(location.Offset + 4)); payload[offset + 4] = (byte)relative; payload[offset + 5] = (byte)(relative >> 8);
                    }
                } else {
                    if (payload.Length < 2 || (payload.Length - 2) % 8 != 0) throw new NotSupportedException("Invalid native ExtSST record.");
                    for (int offset = 2; offset < payload.Length; offset += 8) WriteUInt32(payload, offset, Relocate(BiffRecordReader.ReadUInt32(payload, offset)));
                }
            }
            if (payload.Length > ushort.MaxValue) throw new NotSupportedException("Native record payload exceeds the BIFF record limit.");
            stream.WriteByte((byte)record.Type); stream.WriteByte((byte)(record.Type >> 8));
            stream.WriteByte((byte)payload.Length); stream.WriteByte((byte)(payload.Length >> 8)); stream.Write(payload, 0, payload.Length);
        }
        return stream.ToArray();
    }
    private sealed record OutputRecord(int Key, ushort Type, byte[] Payload);
}
