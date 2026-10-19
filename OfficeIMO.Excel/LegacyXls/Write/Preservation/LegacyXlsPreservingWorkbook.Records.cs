using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservingWorkbook {
    private byte[] RewriteRecords(Dictionary<int, byte[]> replacements, IReadOnlyList<LegacyXlsPreservedString> strings,
        byte[]? filePass, long maximumBytes, CancellationToken cancellationToken, EditingLayout? editing = null) {
        var output = new List<OutputRecord>(_records.Count);
        var locations = _stringLocations.Select(position => (RecordKey: _sstIndex + position.SegmentIndex, position.Offset)).ToList();
        int syntheticKey = -1;
        IReadOnlyList<byte[]> appended = Array.Empty<byte[]>();
        int extKey = _extSstIndex;
        uint stringCount = checked(_stringCount + (uint)strings.Count);
        if (strings.Count != 0) {
            var encoded = LegacyXlsWriter.EncodePreservedStringAppend(strings, maximumBytes, cancellationToken, includeHeader: _sstIndex < 0);
            appended = encoded.Payloads;
            foreach (var position in encoded.Locations) locations.Add((-1 - position.SegmentIndex, position.Offset));
            syntheticKey -= appended.Count;
            if (_sstIndex >= 0) {
                byte[] sst = (byte[])_records[_sstIndex].Payload.Clone(); WriteUInt32(sst, 4, stringCount); replacements[_sstIndex] = sst;
            }
        }
        bool hasSst = _sstIndex >= 0 || strings.Count != 0;
        if (hasSst && extKey < 0) extKey = syntheticKey--;
        int bucketSize = Math.Max(checked((int)(stringCount / 128)) + 1, 8);
        int bucketCount = checked((int)((stringCount + bucketSize - 1) / bucketSize));
        byte[]? newExtSst = !hasSst ? null : new byte[checked(2 + 8 * bucketCount)];
        ushort encodedBucketSize = checked((ushort)bucketSize);
        if (newExtSst != null) { newExtSst[0] = (byte)encodedBucketSize; newExtSst[1] = (byte)(encodedBucketSize >> 8); }
        var worksheetBofs = new HashSet<int>(_model.Worksheets.Select(sheet => sheet.StreamOffset));
        bool needsCalculation = replacements.Count != 0 || editing != null;
        for (int i = 0; i < _records.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (editing != null && editing.Before.TryGetValue(i, out var before)) output.AddRange(before);
            if (_sstIndex < 0 && strings.Count != 0 && _records[i].Type == 0x000a && !_records.Take(i).Any(record => record.Type == 0x000a)) {
                for (int append = 0; append < appended.Count; append++) output.Add(new OutputRecord(-1 - append, append == 0 ? (ushort)0x00fc : (ushort)0x003c, appended[append]));
                output.Add(new OutputRecord(extKey, 0x00ff, newExtSst!));
            }
            if (_sstIndex >= 0 && i == _sstEnd) {
                for (int append = 0; append < appended.Count; append++) output.Add(new OutputRecord(-1 - append, 0x003c, appended[append]));
                if (_extSstIndex < 0) output.Add(new OutputRecord(extKey, 0x00ff, newExtSst!));
            }
            var record = _records[i];
            if (record.Type == (ushort)BiffRecordType.FilePass) continue;
            byte[] payload = newExtSst != null && i == _extSstIndex ? newExtSst : replacements.TryGetValue(i, out var replacement) ? replacement : record.Payload;
            if (editing != null && editing.Replace.TryGetValue(i, out var edited)) output.AddRange(edited);
            else output.Add(new OutputRecord(i, record.Type, payload));
            if (i == 0 && filePass != null) output.Add(new OutputRecord(syntheticKey--, 0x002f, filePass));
            if (needsCalculation && worksheetBofs.Contains(record.Offset) && (editing == null || !editing.Replace.TryGetValue(i, out var replacementRecords) || replacementRecords.Count != 0) && (i + 1 == _records.Count || _records[i + 1].Type != 0x005e))
                output.Add(new OutputRecord(syntheticKey--, 0x005e, new byte[2]));
        }
        if (editing != null && editing.Before.TryGetValue(_records.Count, out var tail)) output.AddRange(tail);
        foreach (var sst in output.Where(record => record.Type == 0x00fc)) {
            // Counts include retained, unused string slots; cstTotal cannot be smaller than cstUnique.
            byte[] data = (byte[])sst.Payload.Clone();
            WriteUInt32(data, 0, Math.Max(stringCount, checked((uint)output.Count(record => record.Type == 0x00fd))));
            output[output.IndexOf(sst)] = sst with { Payload = data };
            break;
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
            if (record.Type == 0x0085) { payload = (byte[])payload.Clone(); WriteUInt32(payload, 0,
                editing != null && editing.PointerTargets.TryGetValue((record.Key, 0), out int target) ? checked((uint)newOffsets[target]) : Relocate(BiffRecordReader.ReadUInt32(payload, 0))); }
            else if (record.Type == 0x020b) {
                if (payload.Length < 16 || (payload.Length - 16) % 4 != 0) throw new NotSupportedException("Invalid native worksheet Index record.");
                payload = (byte[])payload.Clone();
                uint IndexPointer(int offset) => editing != null && editing.PointerTargets.TryGetValue((record.Key, offset), out int target)
                    ? checked((uint)newOffsets[target]) : Relocate(BiffRecordReader.ReadUInt32(payload, offset));
                WriteUInt32(payload, 12, IndexPointer(12));
                if (editing != null && editing.IndexDbCells.TryGetValue(record.Key, out var dbKeys)) {
                    for (int index = 0; index < dbKeys.Count; index++) WriteUInt32(payload, 16 + index * 4, checked((uint)newOffsets[dbKeys[index]]));
                } else for (int offset = 16; offset < payload.Length; offset += 4) WriteUInt32(payload, offset, IndexPointer(offset));
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
