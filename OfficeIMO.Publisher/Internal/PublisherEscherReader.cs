using OfficeIMO.Drawing.Binary;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

// Publisher uses OfficeArt payloads with format-specific container trailers and
// property-pair client anchors. Shared Core readers own FOPT and BLIP decoding.
internal sealed class PublisherEscherReader {
    private readonly PublisherBinaryData _data;
    private readonly PublisherBinaryData? _delay;
    private readonly PublisherParseContext _context;
    private readonly PublisherEscherData _result = new();
    private readonly HashSet<uint> _shapeIds = new();
    internal PublisherEscherReader(PublisherBinaryData data, PublisherBinaryData? delay, PublisherParseContext context) {
        _data = data; _delay = delay; _context = context;
    }
    internal PublisherEscherData Read() {
        foreach (PublisherEscherRecord record in Records(0, _data.Length)) Visit(record, null, 0);
        return _result;
    }
    private void Visit(PublisherEscherRecord record, PublisherGroupSpace? space, int depth) {
        _context.CheckDepth(depth);
        if (record.Kind == 0xF001) { ReadImages(record); return; }
        if (record.Kind == 0xF004) { ReadShape(record, space); return; }
        if ((record.Initial & 15) != 15) return;
        PublisherGroupSpace? childSpace = space;
        foreach (PublisherEscherRecord child in Records(record.Offset, record.End)) {
            if (record.Kind == 0xF003 && child.Kind == 0xF004) {
                PublisherEscherShape? shape = ReadShape(child, childSpace);
                if (shape != null && shape.IsGroup && shape.GroupCoordinates.HasValue && shape.Bounds.HasValue) {
                    childSpace = new PublisherGroupSpace(shape.GroupCoordinates.Value, shape.Bounds.Value);
                    if (shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical) _context.Add("PUB_GROUP_TRANSFORM_UNASSESSED",
                        "Group rotation or mirroring was not applied to the flattened child scene.", OfficeConversionLossKind.Approximation, ShapeLocation(shape.Id));
                }
            } else Visit(child, childSpace, depth + 1);
        }
    }
    private PublisherEscherShape? ReadShape(PublisherEscherRecord record, PublisherGroupSpace? space) {
        IReadOnlyList<PublisherEscherRecord> children = Records(record.Offset, record.End);
        PublisherEscherRecord? fsp = Single(children, 0xF00A), client = Single(children, 0xF011);
        if (fsp == null) throw new InvalidDataException("Publisher shape has no OfficeArt shape descriptor.");
        _data.Range(fsp.Offset, 8, fsp.End);
        uint flags = _data.U32(fsp.Offset + 4);
        IReadOnlyList<OfficeArtProperty> properties = ReadProperties(children);
        uint id = 0;
        if (client != null) {
            Dictionary<ushort, uint> values = ClientValues(client);
            if (!values.TryGetValue(0x6801, out id)) return null;
        }
        var shape = new PublisherEscherShape(id, fsp.Initial >> 4, flags, properties);
        PublisherEscherRecord? anchor = Single(children, 0xF010), childAnchor = Single(children, 0xF00F), coordinates = Single(children, 0xF009);
        if (anchor != null) {
            Dictionary<ushort, uint> values = ClientValues(anchor);
            if (values.TryGetValue(0x2001, out uint x1) && values.TryGetValue(0x2002, out uint y1)
                && values.TryGetValue(0x2003, out uint x2) && values.TryGetValue(0x2004, out uint y2))
                shape.Bounds = new PublisherNativeRectangle(unchecked((int)x1), unchecked((int)y1), unchecked((int)x2), unchecked((int)y2));
        } else if (childAnchor != null) {
            _data.Range(childAnchor.Offset, 16, childAnchor.End);
            var relative = Rectangle(childAnchor.Offset);
            if (space.HasValue) shape.Bounds = space.Value.Resolve(relative);
            else _context.Add("PUB_GROUP_COORDINATES_UNRESOLVED", "A child anchor has no enclosing group coordinate system.", OfficeConversionLossKind.Omission, ShapeLocation(id));
        }
        if (coordinates != null) { _data.Range(coordinates.Offset, 16, coordinates.End); shape.GroupCoordinates = Rectangle(coordinates.Offset); }
        if (id != 0) {
            if (!_shapeIds.Add(id)) throw new InvalidDataException("Duplicate Publisher OfficeArt object identifier.");
            if (_result.Shapes.Count >= _context.Options.Limits.MaxItems) throw new InvalidDataException("Publisher drawing object limit exceeded.");
            _result.Shapes.Add(shape);
        }
        // The root group leader may define a coordinate system without a publication object ID.
        return shape;
    }
    private IReadOnlyList<OfficeArtProperty> ReadProperties(IReadOnlyList<PublisherEscherRecord> children) {
        var result = new List<OfficeArtProperty>();
        foreach (PublisherEscherRecord property in children.Where(item => item.Kind is 0xF00B or 0xF122)) {
            int count = property.Initial >> 4;
            if (count * 6 > property.Length) throw new InvalidDataException("Truncated Publisher OfficeArt property table.");
            var decoded = OfficeArtPropertyTableReader.Read(_data.Bytes, property.Offset, property.Length, (ushort)count);
            foreach (OfficeArtProperty value in decoded) {
                if (value.IsComplex && value.AvailableComplexDataLength != value.Value)
                    throw new InvalidDataException("Truncated Publisher OfficeArt complex property.");
                result.Add(value);
            }
        }
        return result;
    }
    private Dictionary<ushort, uint> ClientValues(PublisherEscherRecord record) {
        _data.Range(record.Offset, 4, record.End);
        if (_data.U32(record.Offset) != record.Length || (record.Length - 4) % 6 != 0)
            throw new InvalidDataException("Invalid Publisher OfficeArt client record length.");
        var result = new Dictionary<ushort, uint>();
        for (int position = record.Offset + 4; position < record.End; position += 6) {
            _context.Record(); ushort id = _data.U16(position);
            if (result.ContainsKey(id)) throw new InvalidDataException("Duplicate Publisher OfficeArt client property.");
            result.Add(id, _data.U32(position + 2));
        }
        return result;
    }
    private void ReadImages(PublisherEscherRecord store) {
        int index = 0;
        foreach (PublisherEscherRecord record in Records(store.Offset, store.End)) {
            if (record.Kind != 0xF007) throw new InvalidDataException("Unexpected Publisher image store record.");
            index++;
            int remaining = (int)(_context.Options.MaximumTotalImageBytes - _context.ImageBytes);
            if (!OfficeArtBlipStoreEntryReader.TryRead(_data.Bytes, record.Offset, record.Length,
                (ushort)(record.Initial >> 4), _delay?.Bytes, out OfficeArtBlipStoreEntry? image,
                Math.Min(_context.Options.MaximumImageBytes, remaining)))
                throw new InvalidDataException("Invalid Publisher image store entry.");
            if (image!.WasImageRejectedBySizeLimit) throw new InvalidDataException("Publisher decoded image byte limit exceeded.");
            if (image.IsPayloadTruncated) throw new InvalidDataException("Truncated Publisher image payload.");
            byte[]? bytes = image.HasImportableImage ? image.ImageBytes : ReadPublisherGif(image, record, Math.Min(_context.Options.MaximumImageBytes, remaining));
            if (bytes != null) {
                _context.AccountImage(bytes.Length);
                _result.Images.Add(index, new PublisherImage(index, image.HasImportableImage ? image.ContentType! : "image/gif", bytes));
            } else _context.Add("PUB_IMAGE_PAYLOAD_UNAVAILABLE", "An image store entry has no recoverable embedded payload; external resources were not fetched.",
                OfficeConversionLossKind.Omission, "Escher/image/" + index);
        }
    }
    private byte[]? ReadPublisherGif(OfficeArtBlipStoreEntry image, PublisherEscherRecord record, int maximumBytes) {
        // Publisher stores GIF bytes in the nominal OfficeArt PNG record. This
        // application-specific envelope differs from the shared MS-ODRAW codec.
        if (image.BlipRecordType != 0xF01E || image.BlipRecordInstance is not (0x6E0 or 0x6E1)) return null;
        PublisherBinaryData data = image.Storage == OfficeArtBlipStorage.Delayed ? _delay! : _data;
        int start = image.Storage == OfficeArtBlipStorage.Delayed
            ? data.Offset(image.DelayedStreamOffset) : checked(record.Offset + 36 + image.NameByteCount);
        int prefix = image.BlipRecordInstance == 0x6E0 ? 17 : 33;
        int size = checked((int)image.BlipPayloadLength!.Value - prefix);
        if (size < 6) return null;
        int position = checked(start + 8 + prefix);
        data.Range(position, size);
        if (data.U32(position) != 0x38464947 || data.U16(position + 4) is not (0x6137 or 0x6139)) return null;
        if (size > maximumBytes) throw new InvalidDataException("Publisher decoded image byte limit exceeded.");
        var bytes = new byte[size];
        Buffer.BlockCopy(data.Bytes, position, bytes, 0, size);
        return bytes;
    }
    private IReadOnlyList<PublisherEscherRecord> Records(int start, int end) {
        _data.Range(start, end - start);
        var result = new List<PublisherEscherRecord>();
        for (int position = start; position < end;) {
            _context.Record(); _data.Range(position, 8, end);
            ushort initial = _data.U16(position), type = _data.U16(position + 2);
            int length = _data.Offset(_data.U32(position + 4));
            _data.Range(position + 8, length, end);
            var record = new PublisherEscherRecord(initial, type, position + 8, length);
            result.Add(record);
            // Publisher stores a drawing-instance separator between top-level
            // containers. The final container ends at EOF without a separator.
            int tail = type is 0xF000 or 0xF002 && record.End < end ? 4 : 0;
            _data.Range(record.End, tail, end);
            position = record.End + tail;
        }
        return result;
    }
    private PublisherNativeRectangle Rectangle(int offset) => new(_data.I32(offset), _data.I32(offset + 4), _data.I32(offset + 8), _data.I32(offset + 12));
    private static PublisherEscherRecord? Single(IReadOnlyList<PublisherEscherRecord> records, int type) {
        var matches = records.Where(item => item.Kind == type).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Ambiguous Publisher OfficeArt shape record.");
        return matches.FirstOrDefault();
    }
    internal static string ShapeLocation(uint id) => "Contents/object/" + id;
}

internal sealed class PublisherEscherRecord {
    internal PublisherEscherRecord(ushort initial, ushort kind, int offset, int length) { Initial = initial; Kind = kind; Offset = offset; Length = length; }
    internal ushort Initial { get; }
    internal ushort Kind { get; }
    internal int Offset { get; }
    internal int Length { get; }
    internal int End => Offset + Length;
}

internal sealed class PublisherEscherData {
    internal List<PublisherEscherShape> Shapes { get; } = new();
    internal Dictionary<int, PublisherImage> Images { get; } = new();
}

internal sealed class PublisherEscherShape {
    internal PublisherEscherShape(uint id, int type, uint flags, IReadOnlyList<OfficeArtProperty> properties) {
        Id = id; Type = type; Flags = flags; Properties = properties;
        Style = OfficeArtShapeStyle.Decode(properties); Transform = OfficeArtShapeTransform.Decode(flags, properties);
    }
    internal uint Id { get; }
    internal int Type { get; }
    internal uint Flags { get; }
    internal bool IsGroup => (Flags & 1) != 0;
    internal IReadOnlyList<OfficeArtProperty> Properties { get; }
    internal OfficeArtShapeStyle Style { get; }
    internal OfficeArtShapeTransform Transform { get; }
    internal PublisherNativeRectangle? Bounds { get; set; }
    internal PublisherNativeRectangle? GroupCoordinates { get; set; }
    internal uint? Property(int id) => Properties.FirstOrDefault(item => item.PropertyId == id)?.Value;
}

internal readonly struct PublisherNativeRectangle {
    internal PublisherNativeRectangle(double x1, double y1, double x2, double y2) { X1 = x1; Y1 = y1; X2 = x2; Y2 = y2; }
    internal double X1 { get; }
    internal double Y1 { get; }
    internal double X2 { get; }
    internal double Y2 { get; }
}
internal readonly struct PublisherGroupSpace {
    internal PublisherGroupSpace(PublisherNativeRectangle coordinates, PublisherNativeRectangle absolute) { Coordinates = coordinates; Absolute = absolute; }
    internal PublisherNativeRectangle Coordinates { get; }
    internal PublisherNativeRectangle Absolute { get; }
    internal PublisherNativeRectangle Resolve(PublisherNativeRectangle relative) {
        double width = Coordinates.X2 - Coordinates.X1, height = Coordinates.Y2 - Coordinates.Y1;
        if (width == 0 || height == 0) throw new InvalidDataException("Publisher group coordinate system is degenerate.");
        double sx = (Absolute.X2 - Absolute.X1) / width, sy = (Absolute.Y2 - Absolute.Y1) / height;
        return new PublisherNativeRectangle(Absolute.X1 + (relative.X1 - Coordinates.X1) * sx, Absolute.Y1 + (relative.Y1 - Coordinates.Y1) * sy,
            Absolute.X1 + (relative.X2 - Coordinates.X1) * sx, Absolute.Y1 + (relative.Y2 - Coordinates.Y1) * sy);
    }
}
