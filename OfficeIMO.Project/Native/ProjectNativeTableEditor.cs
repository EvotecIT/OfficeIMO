using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>Stages mapped record changes without altering the source compound streams.</summary>
internal sealed class ProjectNativeTableEditor : IProjectNativeTableEditor {
    private readonly string _prefix;
    private readonly ProjectNativeTable _table;
    private readonly Dictionary<int, int> _rows = new Dictionary<int, int>();
    private readonly List<byte[]> _metadata = new List<byte[]>(), _secondaryMetadata = new List<byte[]>();
    private readonly MemoryStream _fixed = new MemoryStream(), _secondary = new MemoryStream();
    private readonly Dictionary<int, Dictionary<uint, byte[]>> _variable = new Dictionary<int, Dictionary<uint, byte[]>>();
    private readonly byte[] _fixedHeader, _secondaryHeader, _variableHeader;
    private readonly int _metadataWidth, _secondaryMetadataWidth, _dataWidth, _secondaryDataWidth;
    private readonly long _budget;
    private readonly CancellationToken _token;
    private bool _dirty;
    private readonly ProjectNativeProfile _profile;
    private readonly bool _hasSecondaryStorage;
    private readonly HashSet<int> _deletedRows = new HashSet<int>();

    internal ProjectNativeTableEditor(OfficeCompoundFile source, Dictionary<uint, ProjectNativeValue> properties, string name, uint tableId,
        uint uidField, long budget, CancellationToken token, ProjectNativeProfile profile) {
        _prefix = profile.DataRoot + "/TBknd" + name + "/"; _budget = budget; _token = token; _profile = profile;
        _table = new ProjectNativeTable(source, "TBknd" + name, properties[0x03000000 | tableId],
            properties.TryGetValue(0x00020000 | tableId, out var map) ? map : (ProjectNativeValue?)null, int.MaxValue, token, profile);
        byte[] first = source.Streams[_prefix + "FixedMeta"];
        // Secondary storage is optional in older files. Preserve it when a later
        // producer supplied it, without adding it to files that never had it.
        byte[] second = source.Streams.TryGetValue(_prefix + "Fixed2Meta", out var storedSecond) ? storedSecond : first;
        _hasSecondaryStorage = storedSecond != null;
        _fixedHeader = first.Take(16).ToArray(); _secondaryHeader = second.Take(16).ToArray();
        _variableHeader = source.Streams[_prefix + "VarMeta"].Take(24).ToArray();
        int count = BitConverter.ToInt32(first, 8);
        // The producer reserves a trailing bitmap byte even when the field count is divisible by eight.
        _metadataWidth = _table.MetadataWidth;
        _secondaryMetadataWidth = _hasSecondaryStorage ? _table.SecondaryMetadataWidth : _metadataWidth;
        _dataWidth = _table.Fields.Values.Where(f => f.Source == 10 && !f.Secondary).Select(f => checked(f.Offset + f.Size)).DefaultIfEmpty(0).Max();
        _secondaryDataWidth = _table.Fields.Values.Where(f => f.Source == 10 && f.Secondary).Select(f => checked(f.Offset + f.Size)).DefaultIfEmpty(0).Max();
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested();
            _metadata.Add(Slice(first, 16 + index * _metadataWidth, _metadataWidth));
            _secondaryMetadata.Add(Slice(second, 16 + index * _secondaryMetadataWidth, _secondaryMetadataWidth));
        }
        AppendSource(_fixed, source.Streams[_prefix + "FixedData"]);
        if (source.Streams.TryGetValue(_prefix + "Fixed2Data", out var secondaryData)) AppendSource(_secondary, secondaryData);
        foreach (var record in _table.Records) {
            int uid = record.Integer(uidField) ?? throw new InvalidDataException("Native record UID is absent.");
            if (_rows.ContainsKey(uid)) throw new InvalidDataException("Duplicate native record UID.");
            _rows.Add(uid, record.MetadataIndex);
        }
        ValidateRecordRanges(_metadata, _dataWidth);
        if (_hasSecondaryStorage) ValidateRecordRanges(_secondaryMetadata, _secondaryDataWidth);
        foreach (var row in _table.Variable) _variable.Add(row.Key, row.Value.ToDictionary(pair => pair.Key, pair => pair.Value.Copy()));
    }

    private void ValidateRecordRanges(List<byte[]> metadata, int width) {
        long end = 0;
        foreach (int offset in _rows.Values.Select(index => BitConverter.ToInt32(metadata[index], 4)).OrderBy(value => value)) {
            _token.ThrowIfCancellationRequested();
            if (offset < end) throw new InvalidDataException("Overlapping native record ranges cannot be edited safely.");
            end = (long)offset + width;
        }
    }

    public bool Contains(int uid) => _rows.ContainsKey(uid);
    public bool HasField(uint id) => _table.Fields.ContainsKey(id);
    public IEnumerable<int> Uids => _rows.Keys;
    public void Integer(int uid, uint id, int value) {
        if (!_table.Fields.TryGetValue(id, out var field)) throw new NotSupportedException("Missing native field.");
        Set(uid, id, field.Size == 2 ? BitConverter.GetBytes(checked((short)value)) : BitConverter.GetBytes(value));
    }
    public void Set(int uid, uint id, byte[]? value) {
        _token.ThrowIfCancellationRequested();
        _dirty = true;
        int row = _rows[uid];
        if (!_table.Fields.TryGetValue(id, out var field)) throw new NotSupportedException("Native field is absent from the producer map: " + id.ToString("X8"));
        var flags = field.Secondary ? _secondaryMetadata[row] : _metadata[row];
        if (field.Source == 10) {
            if (value != null && value.Length != field.Size) throw new InvalidDataException("Native fixed value width differs from its mapping.");
            SetPresence(flags, field.Position, value != null);
            if (value != null) {
                var stream = field.Secondary ? _secondary : _fixed;
                int offset = BitConverter.ToInt32(flags, 4);
                stream.Position = checked(offset + field.Offset); stream.Write(value, 0, value.Length);
            }
        } else if (field.Source == 18 || field.Source == 19 || field.Source == 22) {
            SetPresence(flags, field.Position, value != null && value.Any(b => b != 0));
        } else {
            if (value != null && field.Type == 8 && field.Size > 0 && value.Length > field.Size)
                throw new ArgumentException("Native text exceeds this field's declared UTF-16 capacity.", nameof(value));
            if (!_variable.TryGetValue(uid, out var values)) _variable.Add(uid, values = new Dictionary<uint, byte[]>());
            if (value == null) values.Remove(id); else values[id] = (byte[])value.Clone();
            SetPresence(flags, field.Position, value != null);
        }
    }

    public void Add(int uid) {
        _dirty = true;
        if (_rows.ContainsKey(uid)) throw new InvalidOperationException("Native UID already exists.");
        var first = new byte[_metadataWidth]; var second = new byte[_secondaryMetadataWidth];
        Put(first, 4, checked((int)_fixed.Length)); Put(second, 4, checked((int)_secondary.Length));
        Append(_fixed, new byte[_dataWidth]); Append(_secondary, new byte[_secondaryDataWidth]);
        _rows.Add(uid, _metadata.Count); _metadata.Add(first); _secondaryMetadata.Add(second);
    }
    public void AddReserved(int index) {
        _dirty = true;
        var first = new byte[_metadataWidth]; var second = new byte[_secondaryMetadataWidth]; first[0] = 4;
        Put(first, 4, checked((int)_fixed.Length)); Put(second, 4, checked((int)_secondary.Length));
        var sentinel = new byte[_profile.HasExtendedRecords ? 16 : 8]; Put(sentinel, 0, unchecked((int)0xffff0000) + index);
        Append(_fixed, sentinel); Append(_secondary, new byte[_secondaryDataWidth]);
        _metadata.Add(first); _secondaryMetadata.Add(second);
    }

    public void Delete(int uid) {
        _dirty = true;
        int row = _rows[uid]; _deletedRows.Add(row);
        _variable.Remove(uid); _rows.Remove(uid);
    }

    public void Export(Dictionary<string, byte[]> replacements) {
        _token.ThrowIfCancellationRequested();
        if (!_dirty) return;
        using var variable = new MemoryStream(); using var metadata = new MemoryStream();
        Append(metadata, _variableHeader);
        int count = 0;
        foreach (var row in _variable.OrderBy(pair => pair.Key)) {
            _token.ThrowIfCancellationRequested();
            if (_rows.TryGetValue(row.Key, out int index)) {
                int limit = _profile == ProjectNativeProfile.Mpp9 ? byte.MaxValue : ushort.MaxValue;
                if (row.Value.Count > limit) throw new InvalidDataException("Too many native variable values in one record.");
                if (_profile == ProjectNativeProfile.Mpp9) _metadata[index][1] = (byte)row.Value.Count;
                else Put(_metadata[index], 2, (ushort)row.Value.Count);
            }
            foreach (var field in row.Value.OrderBy(pair => pair.Key)) {
                int key = row.Key;
                if (_profile == ProjectNativeProfile.Mpp9) {
                    if (key < 0 || key > 0xffffff) throw new NotSupportedException("MPP9 variable records require a 24-bit nonnegative UID.");
                    key = unchecked((int)((uint)key | ((uint)(_table.Fields[field.Key].Offset >> 16) & 255) << 24));
                }
                Append(metadata, BitConverter.GetBytes(key)); Append(metadata, BitConverter.GetBytes(checked((int)variable.Length)));
                if (_profile != ProjectNativeProfile.Mpp9) Append(metadata, BitConverter.GetBytes(field.Key));
                Append(variable, BitConverter.GetBytes(field.Value.Length)); Append(variable, field.Value); count++;
            }
        }
        byte[] metaBytes = metadata.ToArray(); Put(metaBytes, 8, count); Put(metaBytes, 20, checked((int)variable.Length));
        replacements[_prefix + "VarMeta"] = metaBytes; replacements[_prefix + "Var2Data"] = variable.ToArray();
        var indexes = Enumerable.Range(0, _metadata.Count).Where(i => !_deletedRows.Contains(i)).ToArray();
        replacements[_prefix + "FixedMeta"] = Metadata(_fixedHeader, indexes.Select(i => _metadata[i]).ToArray(), _fixed.Length);
        replacements[_prefix + "FixedData"] = _fixed.ToArray();
        if (_hasSecondaryStorage) {
            replacements[_prefix + "Fixed2Meta"] = Metadata(_secondaryHeader, indexes.Select(i => _secondaryMetadata[i]).ToArray(), _secondary.Length);
            replacements[_prefix + "Fixed2Data"] = _secondary.ToArray();
        }
    }

    private byte[] Metadata(byte[] header, IReadOnlyList<byte[]> rows, long length) {
        using var stream = new MemoryStream(); Append(stream, header); foreach (var row in rows) Append(stream, row);
        byte[] result = stream.ToArray(); Put(result, 8, rows.Count); Put(result, 12, checked((int)length)); return result;
    }
    private void Append(MemoryStream stream, byte[] bytes) {
        _token.ThrowIfCancellationRequested();
        if (bytes.Length > _budget - stream.Length)
            throw OfficeOutputLimit.Create("Native output stream exceeds its byte budget.");
        stream.Position = stream.Length; stream.Write(bytes, 0, bytes.Length);
    }
    private static void AppendSource(MemoryStream stream, byte[] bytes) {
        stream.Position = stream.Length; stream.Write(bytes, 0, bytes.Length);
    }
    private static void SetPresence(byte[] bytes, int position, bool present) {
        int offset = 8 + position / 8;
        if (offset >= bytes.Length) throw new InvalidDataException("Native field bitmap is truncated.");
        byte mask = (byte)(1 << position % 8); bytes[offset] = present ? (byte)(bytes[offset] | mask) : (byte)(bytes[offset] & ~mask);
    }
    private static void Put(byte[] bytes, int offset, int value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 4);
    private static void Put(byte[] bytes, int offset, ushort value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 2);
    private static byte[] Slice(byte[] source, int offset, int count) { var bytes = new byte[count]; Buffer.BlockCopy(source, offset, bytes, 0, count); return bytes; }
    public void Dispose() { _fixed.Dispose(); _secondary.Dispose(); }
}
