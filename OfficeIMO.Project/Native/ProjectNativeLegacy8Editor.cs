using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>Edits Project 98 fixed records and rebuilds their bounded deferred allocation chains.</summary>
internal sealed class ProjectNativeLegacy8Editor : IProjectNativeTableEditor {
    private readonly ProjectNativeTable _table;
    private readonly string _prefix;
    private readonly int _width, _flags, _presence, _indexed;
    private readonly long _budget;
    private readonly CancellationToken _token;
    private readonly List<byte[]> _records = new List<byte[]>();
    private readonly Dictionary<int, int> _rows = new Dictionary<int, int>();
    private readonly HashSet<int> _deleted = new HashSet<int>();
    private readonly Dictionary<int, Dictionary<uint, byte[]>> _values = new Dictionary<int, Dictionary<uint, byte[]>>();
    private bool _dirty;

    internal ProjectNativeLegacy8Editor(OfficeCompoundFile file, Dictionary<uint, ProjectNativeValue> properties, string name,
        uint tableId, uint uidField, long budget, CancellationToken token) {
        _prefix = ProjectNativeProfile.Mpp8.DataRoot + "/TBknd" + name + "/"; _budget = budget; _token = token;
        var descriptor = properties[0x02000000 | (tableId - 0x13)];
        var layout = ProjectNativeProperties.Read(descriptor.Copy(), token);
        _width = layout[5].Int32(); _flags = layout[6].Int32(); _presence = layout[7].Int32(); _indexed = layout[8].Int32();
        _table = new ProjectNativeTable(file, "TBknd" + name, descriptor, int.MaxValue, token);
        byte[] source = file.Streams[_prefix + "FixFix   0"];
        for (int at = 0; at < source.Length; at += _width) {
            token.ThrowIfCancellationRequested();
            _records.Add(new ProjectNativeValue(source, at, _width).Copy());
        }
        foreach (var record in _table.Records) {
            int uid = record.Integer(uidField) ?? throw new InvalidDataException("Missing Project 98 UID.");
            if (_rows.ContainsKey(uid)) throw new InvalidDataException("Duplicate Project 98 UID.");
            _rows.Add(uid, record.MetadataIndex);
        }
        foreach (var row in _table.Variable) _values.Add(row.Key, row.Value.ToDictionary(v => v.Key, v => v.Value.Copy()));
    }

    public bool Contains(int uid) => _rows.ContainsKey(uid);
    public bool HasField(uint id) => _table.Fields.ContainsKey(id);
    public IEnumerable<int> Uids => _rows.Keys;
    public void Integer(int uid, uint id, int value) {
        if (!_table.Fields.TryGetValue(id, out var field)) throw new NotSupportedException("Missing Project 98 integer field.");
        Set(uid, id, field.Size == 2 ? BitConverter.GetBytes(checked((short)value)) : BitConverter.GetBytes(value));
    }

    public void Set(int uid, uint id, byte[]? value) {
        _token.ThrowIfCancellationRequested();
        if (!_table.Fields.TryGetValue(id, out var field)) throw new NotSupportedException("Missing Project 98 mapped field: " + id.ToString("X8"));
        var row = _records[_rows[uid]];
        if (field.Source == 10) {
            if (value != null) {
                value = EncodeNumber(field, value);
                if (value.Length != field.Size || field.Offset > row.Length - value.Length) throw new ArgumentException("Project 98 fixed value width differs from its descriptor.");
                Buffer.BlockCopy(value, 0, row, field.Offset, value.Length);
            }
        } else if (field.Source == 18) {
            uint bits = new ProjectNativeValue(row, 0, row.Length).UInt32(field.Offset);
            bits = value != null && value.Any(b => b != 0) ? bits | field.Mask : bits & ~field.Mask;
            Put(row, field.Offset, unchecked((int)bits));
        } else {
            if (value != null && field.Type == 8 && field.Size > 0 && value.Length > field.Size) throw new ArgumentException("Project 98 text exceeds its declared UTF-16 capacity.");
            if (!_values.TryGetValue(uid, out var values)) _values.Add(uid, values = new Dictionary<uint, byte[]>());
            if (value == null) values.Remove(id); else values[id] = EncodeNumber(field, value);
        }
        Presence(row, field, value != null); _dirty = true;
    }

    private static byte[] EncodeNumber(ProjectNativeField field, byte[] value) {
        if (value.Length == 8 && (field.Type == 101 || field.Type == 102)) {
            double number = BitConverter.ToDouble(value, 0);
            double minimum = field.Type == 101 ? -140737488355328d : int.MinValue;
            double maximum = field.Type == 101 ? 140737488355327d : int.MaxValue;
            if (double.IsNaN(number) || double.IsInfinity(number) || number != Math.Truncate(number) || number < minimum || number > maximum)
                throw new ArgumentOutOfRangeException(nameof(value), "Project 98 scaled numeric precision or range would be lost.");
            return BitConverter.GetBytes((long)number).Take(field.Type == 101 ? 6 : 4).ToArray();
        }
        return (byte[])value.Clone();
    }

    public void Add(int uid) {
        _token.ThrowIfCancellationRequested();
        if (_rows.ContainsKey(uid)) throw new InvalidOperationException("Project 98 UID already exists.");
        if ((long)(_records.Count + 1) * _width > _budget) throw OfficeOutputLimit.Create("Project 98 fixed table exceeds its byte budget.");
        _rows.Add(uid, _records.Count); _records.Add(new byte[_width]); _dirty = true;
    }
    public void AddReserved(int index) {
        if ((long)(_records.Count + 1) * _width > _budget) throw OfficeOutputLimit.Create("Project 98 reserved table exceeds its byte budget.");
        var row = new byte[_width]; Put(row, 0, unchecked((int)0xffff0000) + index); Put(row, _flags, 4);
        _records.Add(row); _dirty = true;
    }
    public void Delete(int uid) { _deleted.Add(_rows[uid]); _rows.Remove(uid); _values.Remove(uid); _dirty = true; }

    public void Export(Dictionary<string, byte[]> replacements) {
        _token.ThrowIfCancellationRequested(); if (!_dirty) return;
        using var variable = new OfficeBoundedMemoryStream(_budget); using var writer = new BinaryWriter(variable);
        writer.Write(-1);
        int Allocate(byte[] value) {
            int start = checked((int)variable.Length), count = Math.Max(1, checked(value.Length + 35) / 32), written = 0;
            for (int index = 0; index < count; index++) {
                _token.ThrowIfCancellationRequested();
                var block = new byte[36]; int header = index == 0 ? 8 : 4;
                Put(block, 0, index == count - 1 ? -1 : checked(start + (index + 1) * 36));
                if (index == 0) Put(block, 4, value.Length);
                int take = Math.Min(36 - header, value.Length - written);
                Buffer.BlockCopy(value, written, block, header, take); written += take; writer.Write(block);
            }
            return ~start;
        }
        foreach (var rowPair in _rows.OrderBy(p => p.Value)) {
            _token.ThrowIfCancellationRequested(); var row = _records[rowPair.Value];
            using var indexed = new OfficeBoundedMemoryStream(_budget); using var entries = new BinaryWriter(indexed);
            if (_values.TryGetValue(rowPair.Key, out var values)) foreach (var pair in values.OrderBy(p => _table.Fields[p.Key].Position)) {
                var field = _table.Fields[pair.Key];
                if (field.Offset != field.Position) Put(row, field.Offset, Allocate(pair.Value));
                else {
                    if (_indexed < 0) throw new NotSupportedException("This Project 98 table has no indexed field storage.");
                    byte[] bytes = field.LegacyIndirect ? BitConverter.GetBytes(Allocate(pair.Value)) : pair.Value;
                    entries.Write(bytes.Length); entries.Write(field.Position); entries.Write(bytes);
                    if ((bytes.Length & 1) != 0) entries.Write((byte)0);
                }
            }
            if (_indexed >= 0) Put(row, _indexed, indexed.Length == 0 ? 0 : Allocate(indexed.ToArray()));
        }
        using var fixedOutput = new OfficeBoundedMemoryStream(_budget);
        for (int index = 0; index < _records.Count; index++) {
            _token.ThrowIfCancellationRequested();
            if (_deleted.Contains(index)) continue;
            var bytes = _records[index]; fixedOutput.Write(bytes, 0, bytes.Length);
        }
        replacements[_prefix + "FixFix   0"] = fixedOutput.ToArray();
        if (_indexed >= 0 || _values.Count > 0) replacements[_prefix + "FixDeferFix   0"] = variable.ToArray();
    }
    private void Presence(byte[] row, ProjectNativeField field, bool present) {
        int offset = checked(_presence + field.Position / 8); byte mask = (byte)(1 << field.Position % 8);
        row[offset] = present ? (byte)(row[offset] | mask) : (byte)(row[offset] & ~mask);
    }
    private static void Put(byte[] bytes, int offset, int value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, offset, 4);
    public void Dispose() { }
}
