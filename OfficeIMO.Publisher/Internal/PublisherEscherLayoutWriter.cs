using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Grows selected local FOPT tables and their containing record lengths, retaining native separators and opaque records.</summary>
internal sealed class PublisherEscherLayoutWriter {
    private readonly PublisherBinaryData _data;
    private readonly PublisherParseContext _context;
    private readonly PublisherEscherReader _reader;
    private readonly IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, uint>> _changes;
    private readonly List<PublisherNativeStreamEdit> _edits = new();
    private int _seen;

    private PublisherEscherLayoutWriter(byte[] source, IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, uint>> changes,
        PublisherParseContext context) {
        _data = new(source, "Escher/EscherStm"); _context = context;
        _reader = new(_data, null, context); _changes = changes;
    }

    internal static byte[]? Rewrite(byte[] source, IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, uint>> changes,
        PublisherParseContext context, long maximumBytes) {
        var writer = new PublisherEscherLayoutWriter(source, changes, context);
        if (!writer.Visit(0, source.Length, 0, out _) || writer._seen != changes.Count) return null;
        return PublisherNativeStreamEdit.Apply(source, writer._edits, maximumBytes, context.Token);
    }

    private bool Visit(int start, int end, int depth, out int growth) {
        _context.CheckDepth(depth); growth = 0;
        foreach (PublisherEscherRecord record in _reader.Records(start, end)) {
            int delta = 0;
            if (record.Kind == 0xF004) {
                if (_changes.TryGetValue(record.Offset, out var changes)) {
                    _seen++;
                    if (!Shape(record, changes, out delta)) return false;
                }
            } else if ((record.Initial & 15) == 15 && record.Kind != 0xF001) {
                if (!Visit(record.Offset, record.End, depth + 1, out delta)) return false;
            }
            if (delta != 0) _edits.Add(new(record.Offset - 4, 4, BitConverter.GetBytes(checked((uint)(record.Length + delta)))));
            growth = checked(growth + delta);
        }
        return true;
    }

    private bool Shape(PublisherEscherRecord shape, IReadOnlyDictionary<ushort, uint> changes, out int growth) {
        growth = 0;
        var tables = _reader.Records(shape.Offset, shape.End).Where(record => record.Kind is 0xF00B or 0xF122).ToArray();
        var owners = new Dictionary<ushort, PublisherEscherRecord>();
        foreach (PublisherEscherRecord table in tables) {
            if ((table.Initial & 15) != 3 || (table.Initial >> 4) * 6 > table.Length) return false;
            for (int index = 0; index < table.Initial >> 4; index++) {
                _context.Record(); ushort id = (ushort)(_data.U16(table.Offset + index * 6) & 0x3FFF);
                if (!changes.ContainsKey(id)) continue;
                if (owners.ContainsKey(id)) return false;
                owners.Add(id, table);
            }
        }
        var primary = tables.Where(table => table.Kind == 0xF00B).ToArray();
        foreach (ushort id in changes.Keys.Where(id => !owners.ContainsKey(id))) {
            if (primary.Length != 1) return false;
            owners.Add(id, primary[0]);
        }
        foreach (var group in owners.GroupBy(pair => pair.Value)) {
            PublisherEscherRecord table = group.Key;
            byte[]? body = OfficeArtPropertyTableEditor.RewriteScalars(_data.Bytes, table.Offset, table.Length,
                (ushort)(table.Initial >> 4), group.ToDictionary(pair => pair.Key, pair => changes[pair.Key]), out ushort count);
            if (body == null) return false;
            byte[] record = new byte[checked(body.Length + 8)];
            BitConverter.GetBytes((ushort)((count << 4) | 3)).CopyTo(record, 0);
            BitConverter.GetBytes(table.Kind).CopyTo(record, 2);
            BitConverter.GetBytes((uint)body.Length).CopyTo(record, 4); body.CopyTo(record, 8);
            _edits.Add(new(table.Offset - 8, table.Length + 8, record));
            growth = checked(growth + body.Length - table.Length);
        }
        return true;
    }
}
