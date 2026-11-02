using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Composes source-bound image-store and local FOPT changes with enclosing lengths, preserving separators and opaque records.</summary>
internal sealed class PublisherEscherSourceWriter {
    private readonly PublisherBinaryData _data;
    private readonly PublisherParseContext _context;
    private readonly PublisherEscherReader _reader;
    private readonly IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>> _changes;
    private readonly PublisherNativeStreamEdit? _imageStore;
    private readonly List<PublisherNativeStreamEdit> _edits = new();
    private int _seen;
    private bool _imageStoreSeen;
    private readonly PublisherEscherShapeEdits? _shapes;
    private int _removedSeen, _appendedSeen;

    private PublisherEscherSourceWriter(byte[] source, IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>> changes,
        PublisherNativeStreamEdit? imageStore, PublisherParseContext context, PublisherEscherShapeEdits? shapes) {
        _data = new(source, "Escher/EscherStm"); _context = context;
        _reader = new(_data, null, context); _changes = changes; _imageStore = imageStore;
        _shapes = shapes;
    }

    internal static byte[]? Rewrite(byte[] source, IReadOnlyDictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>> changes,
        PublisherNativeStreamEdit? imageStore, PublisherParseContext context, long maximumBytes, PublisherEscherShapeEdits? shapes = null) {
        var writer = new PublisherEscherSourceWriter(source, changes, imageStore, context, shapes);
        if (!writer.Visit(0, source.Length, 0, out _) || writer._seen != changes.Count || imageStore != null && !writer._imageStoreSeen) return null;
        if (shapes != null && (writer._removedSeen != shapes.Removed.Count || writer._appendedSeen != shapes.Appended.Count)) return null;
        return PublisherNativeStreamEdit.Apply(source, writer._edits, maximumBytes, context.Token);
    }

    private bool Visit(int start, int end, int depth, out int growth) {
        _context.CheckDepth(depth); growth = 0;
        foreach (PublisherEscherRecord record in _reader.Records(start, end)) {
            int delta = 0;
            bool replaced = false;
            if (record.Kind == 0xF001 && _imageStore != null && record.Offset - 8 == _imageStore.Offset) {
                if (_imageStore.Length != record.Length + 8 || _imageStoreSeen) return false;
                _imageStoreSeen = true; _edits.Add(_imageStore); replaced = true;
                delta = _imageStore.Bytes.Length - _imageStore.Length;
            } else if (record.Kind == 0xF004) {
                if (_shapes?.Removed.Contains(record.Offset) == true) {
                    if (_changes.ContainsKey(record.Offset)) return false;
                    _removedSeen++; replaced = true;
                    _edits.Add(new(record.Offset - 8, record.Length + 8, Array.Empty<byte>()));
                    delta = -record.Length - 8;
                }
                else if (_changes.TryGetValue(record.Offset, out var changes)) {
                    _seen++;
                    if (!Shape(record, changes, out delta)) return false;
                }
            } else if ((record.Initial & 15) == 15 && record.Kind != 0xF001) {
                if (!Visit(record.Offset, record.End, depth + 1, out delta)) return false;
            }
            if (record.Kind == 0xF003 && _shapes != null && _shapes.Appended.TryGetValue(record.Offset, out byte[]? appended)) {
                _appendedSeen++; _edits.Add(new(record.End, 0, appended)); delta = checked(delta + appended.Length);
            }
            if (delta != 0 && !replaced) _edits.Add(new(record.Offset - 4, 4, BitConverter.GetBytes(checked((uint)(record.Length + delta)))));
            growth = checked(growth + delta);
        }
        return true;
    }

    private bool Shape(PublisherEscherRecord shape, IReadOnlyDictionary<ushort, OfficeArtProperty> changes, out int growth) {
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
            var selected = group.ToDictionary(pair => pair.Key, pair => changes[pair.Key]);
            byte[]? body; ushort count;
            if (selected.Values.All(property => !property.IsComplex && !property.IsBlipId))
                body = OfficeArtPropertyTableEditor.RewriteScalars(_data.Bytes, table.Offset, table.Length,
                    (ushort)(table.Initial >> 4), selected.ToDictionary(pair => pair.Key, pair => pair.Value.Value), out count);
            else body = OfficeArtPropertyTableEditor.Rewrite(_data.Bytes, table.Offset, table.Length,
                (ushort)(table.Initial >> 4), selected, out count);
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
