using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

/// <summary>Keeps native records/reference tables and edits existing fixed-width literal cell records.</summary>
internal sealed partial class LegacyXlsPreservingWorkbook {
    private readonly OfficeCompoundFile _compound;
    private readonly string _streamName;
    private readonly IReadOnlyList<BiffRecord> _records;
    private readonly Dictionary<(string, string), CellSlot> _cells = new();
    private readonly int _sstIndex;
    private readonly int _sstEnd;
    private readonly int _extSstIndex;
    private readonly uint _stringCount;
    private readonly List<(int SegmentIndex, int Offset)> _stringLocations = new();
    private readonly Dictionary<string, uint> _plainStrings = new(StringComparer.Ordinal);

    internal LegacyXlsPreservingWorkbook(LegacyXlsWorkbook workbook, CancellationToken cancellationToken) {
        _compound = workbook.SourceCompoundFile!;
        byte[] raw = LegacyWorkbookStreamLocator.FindWorkbookStream(_compound.Streams) ?? throw new NotSupportedException("Missing native workbook stream.");
        _streamName = _compound.Streams.First(item => ReferenceEquals(item.Value, raw)).Key;
        if (raw.Length < 8 || BiffRecordReader.ReadUInt16(raw, 0) != 0x0809 || BiffRecordReader.ReadUInt16(raw, 4) != 0x0600)
            throw new NotSupportedException("Native record preservation requires BIFF8.");
        var diagnostics = new List<LegacyXlsImportDiagnostic>();
        _records = BiffRecordReader.ReadRecords(_compound.Streams[_streamName], diagnostics, cancellationToken);
        if (diagnostics.Count != 0 || _records.Count == 0 || _records[0].Type != 0x0809 || BiffRecordReader.ReadUInt16(_records[0].Payload, 0) != 0x0600)
            throw new NotSupportedException("Native record preservation requires a complete BIFF8 workbook stream.");
        _sstIndex = -1; _extSstIndex = -1; _sstEnd = -1;
        for (int i = 1; i < _records.Count && _records[i].Type != 0x000a; i++) {
            if (_records[i].Type == 0x00fc) { if (_sstIndex >= 0) throw new NotSupportedException("Duplicate native SST records."); _sstIndex = i; _sstEnd = i + 1; while (_sstEnd < _records.Count && _records[_sstEnd].Type == 0x003c) _sstEnd++; }
            if (_records[i].Type == 0x00ff) _extSstIndex = i;
        }
        if (_sstIndex >= 0) {
            if (_records[_sstIndex].Payload.Length < 8) throw new NotSupportedException("Native SST preservation requires a complete header.");
            _stringCount = BiffRecordReader.ReadUInt32(_records[_sstIndex].Payload, 4);
            var payloads = _records.Skip(_sstIndex).Take(_sstEnd - _sstIndex).Select(item => item.Payload).ToArray();
            var values = BiffStringReader.ReadSharedStringValues(payloads, diagnostics, _records[_sstIndex].Offset, _stringLocations, cancellationToken);
            if (diagnostics.Count != 0 || values.Count != _stringCount) throw new NotSupportedException("The native SST could not be indexed without loss.");
            for (int index = 0; index < values.Count; index++) {
                var location = _stringLocations[index]; int segment = location.SegmentIndex, offset = location.Offset + 2;
                while (segment < payloads.Length && offset >= payloads[segment].Length) offset -= payloads[segment++].Length;
                if (segment >= payloads.Length) throw new NotSupportedException("Invalid native string header.");
                if ((payloads[segment][offset] & 0x0c) == 0 && !_plainStrings.ContainsKey(values[index].Text))
                    _plainStrings.Add(values[index].Text, checked((uint)index));
            }
        }
        var recordIndexes = _records.Select((record, index) => (record.Offset, Index: index)).ToDictionary(item => item.Offset, item => item.Index);
        foreach (var sheet in workbook.Worksheets) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!recordIndexes.TryGetValue(sheet.StreamOffset, out int firstRecord)) throw new NotSupportedException("The source worksheet BOF is not aligned to a BIFF record.");
            int depth = 0; bool started = false;
            for (int i = firstRecord; i < _records.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                var record = _records[i];
                if (record.Type == 0x0809) { depth++; started = true; continue; }
                if (record.Type == 0x000a) { if (--depth == 0) break; continue; }
                if (!started || depth != 1) continue;
                byte[] data = record.Payload;
                if (record.Type == 0x00bd && data.Length >= 12 && (data.Length - 6) % 6 == 0) {
                    int first = BiffRecordReader.ReadUInt16(data, 2), count = (data.Length - 6) / 6;
                    for (int cell = 0; cell < count; cell++) AddCell(sheet.Name, i, first + cell, 6 + cell * 6, "rk");
                } else if ((record.Type == 0x0203 && data.Length == 14) || (record.Type == 0x027e && data.Length == 10)
                    || (record.Type == 0x00fd && data.Length == 10) || (record.Type == 0x0205 && data.Length == 8)) {
                    string kind = record.Type == 0x0203 ? "number" : record.Type == 0x027e ? "rk" : record.Type == 0x00fd ? "text" : data[7] == 0 ? "boolean" : "error";
                    AddCell(sheet.Name, i, BiffRecordReader.ReadUInt16(data, 2), 6, kind);
                }
            }
        }
    }

    private void AddCell(string sheet, int index, int column, int valueOffset, string kind) {
        int row = BiffRecordReader.ReadUInt16(_records[index].Payload, 0);
        var key = (sheet, A1.CellReference(row + 1, column + 1));
        if (_cells.ContainsKey(key)) throw new NotSupportedException("The preserving profile cannot edit duplicate native cell records.");
        _cells.Add(key, new CellSlot(index, valueOffset, kind));
    }

    internal bool CanApply(LegacyXlsLiteralEdit edit) {
        if (!_cells.TryGetValue((edit.Sheet, edit.Cell), out var slot)) return false;
        if (slot.Kind == "text") return edit.Kind == "text" && edit.Text.Length <= 32767 && _sstIndex >= 0;
        if (slot.Kind == "boolean") return edit.Kind == "boolean" && (edit.Text == "1" || edit.Text == "0" || bool.TryParse(edit.Text, out _));
        if (slot.Kind == "error") return edit.Kind == "error" && TryGetErrorCode(edit.Text, out _);
        return edit.Kind == "number" && double.TryParse(edit.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            && !double.IsNaN(value) && !double.IsInfinity(value) && (slot.Kind != "rk" || TryEncodeRk(value, out _));
    }

    internal byte[] Write(IReadOnlyList<LegacyXlsLiteralEdit> changes, byte[]? filePass, long maximumBytes, CancellationToken cancellationToken, Func<byte[], byte[]>? transform = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_compound.Streams[_streamName].LongLength > maximumBytes) throw OfficeOutputLimit.Create("Native XLS preservation exceeds the output materialization limit.");
        var replacements = new Dictionary<int, byte[]>();
        var strings = new List<string>(); var stringIndexes = new Dictionary<string, uint>(_plainStrings, StringComparer.Ordinal);
        foreach (var edit in changes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!CanApply(edit)) throw new NotSupportedException("The native XLS literal edit is outside the preserving profile.");
            var slot = _cells[(edit.Sheet, edit.Cell)];
            if (!replacements.TryGetValue(slot.Record, out var data)) replacements.Add(slot.Record, data = (byte[])_records[slot.Record].Payload.Clone());
            if (slot.Kind == "text") {
                if (!stringIndexes.TryGetValue(edit.Text, out uint index)) { index = checked(_stringCount + (uint)strings.Count); stringIndexes.Add(edit.Text, index); strings.Add(edit.Text); }
                WriteUInt32(data, slot.ValueOffset, index);
            } else if (slot.Kind == "boolean") data[slot.ValueOffset] = edit.Text == "1" || bool.TryParse(edit.Text, out bool flag) && flag ? (byte)1 : (byte)0;
            else if (slot.Kind == "error") { TryGetErrorCode(edit.Text, out byte error); data[slot.ValueOffset] = error; }
            else {
                double value = double.Parse(edit.Text, CultureInfo.InvariantCulture);
                if (slot.Kind == "rk") { TryEncodeRk(value, out uint rk); WriteUInt32(data, slot.ValueOffset, rk); }
                else Buffer.BlockCopy(BitConverter.GetBytes(value), 0, data, slot.ValueOffset, 8);
            }
        }
        byte[] workbook = RewriteRecords(replacements, strings, filePass, maximumBytes, cancellationToken);
        if (transform != null) workbook = transform(workbook);
        return OfficeCompoundFileWriter.Rewrite(_compound, new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) { [_streamName] = workbook }, maxOutputBytes: maximumBytes);
    }

    private static bool TryEncodeRk(double value, out uint encoded) {
        for (int scaled = 0; scaled <= 1; scaled++) {
            double candidate = scaled == 0 ? value : value * 100;
            if (candidate >= -536870912 && candidate <= 536870911 && Math.Truncate(candidate) == candidate) {
                uint code = unchecked((uint)(int)candidate << 2) | 2U | (uint)scaled;
                if ((scaled == 0 ? (int)code >> 2 : ((int)code >> 2) / 100D) == value) { encoded = code; return true; }
            }
            long bits = BitConverter.DoubleToInt64Bits(candidate);
            if ((bits & 0x3ffffffffL) == 0) {
                uint code = (uint)(bits >> 32) & 0xfffffffcU;
                double decoded = BitConverter.Int64BitsToDouble((long)code << 32);
                if ((scaled == 0 ? decoded : decoded / 100D) == value) { encoded = code | (uint)scaled; return true; }
            }
        }
        encoded = 0; return false;
    }

    private static bool TryGetErrorCode(string text, out byte code) {
        string[] values = { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA" };
        byte[] codes = { 0, 7, 15, 23, 29, 36, 42, 43 };
        int index = Array.IndexOf(values, text); code = index < 0 ? (byte)0 : codes[index]; return index >= 0;
    }
    private static void WriteUInt32(byte[] bytes, int offset, uint value) { for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8)); }
    private sealed record CellSlot(int Record, int ValueOffset, string Kind);
}
