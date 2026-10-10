using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherContentsReader {
    private readonly PublisherBinaryData _data;
    private readonly PublisherBlockReader _blocks;
    private readonly PublisherParseContext _context;
    internal PublisherContentsReader(PublisherBinaryData data, PublisherParseContext context) {
        _data = data; _blocks = new PublisherBlockReader(data, context); _context = context;
    }
    internal PublisherContents Read() {
        int trailer = _data.Offset(_data.U32(0x1A));
        var sections = _blocks.Chunk(trailer);
        PublisherBlock[] directories = sections.Where(item => item.Type == 0x90).ToArray();
        if (directories.Length != 1) throw new InvalidDataException("Publisher chunk directory is missing or ambiguous.");
        var chunks = new Dictionary<uint, PublisherContentChunk>();
        uint id = 0;
        foreach (PublisherBlock slot in _blocks.Children(directories[0])) {
            if (slot.Type == 0x88) {
                var fields = _blocks.Children(slot);
                PublisherBlock? type = Field(fields, 2), offset = Field(fields, 4), parent = Field(fields, 5);
                if (type.HasValue && offset.HasValue) {
                    int start = _data.Offset(offset.Value.Value);
                    if (start < 0x20 || start >= trailer) throw new InvalidDataException("Publisher content chunk overlaps its header or directory.");
                    chunks.Add(id, new PublisherContentChunk(id, type.Value.Value, start, parent?.Value));
                }
            }
            id++;
        }
        PublisherContentChunk[] documents = chunks.Values.Where(item => item.Kind == 0x44).ToArray();
        var orderedChunks = chunks.Values.OrderBy(item => item.Offset).ToArray();
        for (int i = 0; i < orderedChunks.Length; i++) {
            orderedChunks[i].End = i + 1 < orderedChunks.Length ? orderedChunks[i + 1].Offset : trailer;
            if (orderedChunks[i].End <= orderedChunks[i].Offset) throw new InvalidDataException("Overlapping Publisher chunk offsets.");
        }
        if (documents.Length != 1) throw new InvalidDataException("Publisher document descriptor is missing or ambiguous.");
        var document = _blocks.Chunk(documents[0].Offset, documents[0].End);
        PublisherBlock size = Field(document, 0x12) ?? throw new InvalidDataException("Publisher page dimensions are missing.");
        var dimensions = _blocks.Children(size);
        uint width = (Field(dimensions, 1) ?? throw new InvalidDataException("Publisher page width is missing.")).Value;
        uint height = (Field(dimensions, 2) ?? throw new InvalidDataException("Publisher page height is missing.")).Value;
        if (width == 0 || height == 0 || width > int.MaxValue || height > int.MaxValue) throw new InvalidDataException("Invalid Publisher page dimensions.");
        var result = new PublisherContents(width / 12700D, height / 12700D, chunks);
        PublisherBlock pageList = Field(document, 2) ?? throw new InvalidDataException("Publisher page order is missing.");
        var seenPages = new HashSet<uint>();
        foreach (PublisherBlock reference in _blocks.Children(pageList)) {
            if (reference.Type != 0x70) continue;
            uint pageId = reference.Value;
            if (!seenPages.Add(pageId)) throw new InvalidDataException("Duplicate Publisher page reference.");
            if (!chunks.TryGetValue(pageId, out PublisherContentChunk? pageChunk))
                throw new InvalidDataException("Publisher page reference does not resolve to a page descriptor.");
            // Empty reference slots also occur in the native page list.
            if (pageChunk.Kind == 0x59) continue;
            if (pageChunk.Kind != 0x43) throw new InvalidDataException("Publisher page reference has an invalid descriptor kind.");
            var fields = _blocks.Chunk(pageChunk.Offset, pageChunk.End);
            PublisherBlock? nameField = Field(fields, 0x0E);
            string name = nameField.HasValue && nameField.Value.Type == 0xC0
                ? _data.Utf16(nameField.Value.Offset + 4, nameField.Value.Length - 4).TrimEnd('\0') : string.Empty;
            // Printable pages carry a content-count field; named definitions
            // are masters. Remaining unnamed definitions are utility pages.
            if (string.IsNullOrEmpty(name) && !Field(fields, 1).HasValue) continue;
            if (result.Pages.Count >= _context.Options.MaximumPages) throw new InvalidDataException("Publisher page limit exceeded.");
            var page = new PublisherSourcePage(pageId, name, Field(fields, 0x0D)?.Value, Field(fields, 0x0A)?.Value);
            result.Pages.Add(page);
            PublisherBlock? shapes = Field(fields, 2);
            if (shapes.HasValue) {
                foreach (PublisherBlock shape in _blocks.Children(shapes.Value)) {
                    if (shape.Type != 0x70) continue;
                    if (result.ObjectPages.ContainsKey(shape.Value)) throw new InvalidDataException("Publisher object belongs to more than one page.");
                    result.ObjectPages.Add(shape.Value, pageId);
                }
            }
        }
        foreach (PublisherContentChunk chunk in chunks.Values) {
            if (chunk.IsShape) {
                if (result.Shapes.Count >= _context.Options.Limits.MaxItems) throw new InvalidDataException("Publisher object limit exceeded.");
                var shape = new PublisherSourceShape(chunk, _blocks.Chunk(chunk.Offset, chunk.End));
                if (chunk.Kind == 0x10) shape.Table = ReadTable(shape, chunks);
                else if (shape.Value(0x27).HasValue) shape.WrapObjects = ReadWrapReferences(shape);
                result.Shapes.Add(chunk.Id, shape);
            } else if (chunk.Kind == 0x5C) ReadPalette(chunk, result.Palette, chunk.End);
        }
        return result;
    }
    private void ReadPalette(PublisherContentChunk chunk, List<OfficeColor?> colors, int boundary) {
        foreach (PublisherBlock block in _blocks.Chunk(chunk.Offset, boundary)) {
            if (block.Type != 0xA0) continue;
            foreach (PublisherBlock entry in _blocks.Children(block)) {
                // The empty palette slot is Publisher's automatic black.
                OfficeColor? color = entry.Type == 0x78 ? OfficeColor.Black : null;
                if (entry.Type == 0x88) {
                    PublisherBlock? rgb = Field(_blocks.Children(entry), 1);
                    if (rgb.HasValue) color = OfficeColor.FromRgb((byte)rgb.Value.Value, (byte)(rgb.Value.Value >> 8), (byte)(rgb.Value.Value >> 16));
                }
                colors.Add(color);
            }
        }
    }
    private IReadOnlyList<uint> ReadWrapReferences(PublisherSourceShape shape) {
        PublisherBlock? references = Field(shape.Fields, 0x47);
        uint count = shape.Value(0x46) ?? 0;
        if (!references.HasValue) {
            if (count != 0) throw new InvalidDataException("Publisher text-wrap count has no reference list.");
            return Array.Empty<uint>();
        }
        var result = new List<uint>();
        var visited = new HashSet<uint>();
        foreach (PublisherBlock item in _blocks.Children(references.Value)) {
            if (item.Type != 0x68 || item.Id != 0 || !visited.Add(item.Value) || item.Value == shape.Chunk.Id)
                throw new InvalidDataException("Invalid or duplicate Publisher text-wrap reference.");
            if (result.Count >= _context.Options.Limits.MaxItems) throw new InvalidDataException("Publisher text-wrap item limit exceeded.");
            result.Add(item.Value);
        }
        if (result.Count != count) throw new InvalidDataException("Publisher text-wrap count disagrees with its reference list.");
        return result;
    }
    internal static PublisherBlock? Field(IReadOnlyList<PublisherBlock> blocks, byte id) {
        PublisherBlock? result = null;
        foreach (PublisherBlock block in blocks) {
            if (block.Id != id) continue;
            if (result.HasValue) throw new InvalidDataException($"Duplicate Publisher field 0x{id:X2}.");
            result = block;
        }
        return result;
    }
}

internal sealed class PublisherContents {
    internal PublisherContents(double width, double height, Dictionary<uint, PublisherContentChunk> chunks) { Width = width; Height = height; Chunks = chunks; }
    internal double Width { get; }
    internal double Height { get; }
    internal Dictionary<uint, PublisherContentChunk> Chunks { get; }
    internal List<PublisherSourcePage> Pages { get; } = new();
    internal Dictionary<uint, PublisherSourceShape> Shapes { get; } = new();
    internal Dictionary<uint, uint> ObjectPages { get; } = new();
    internal List<OfficeColor?> Palette { get; } = new();
}

internal sealed class PublisherContentChunk {
    internal PublisherContentChunk(uint id, uint kind, int offset, uint? parent) { Id = id; Kind = kind; Offset = offset; Parent = parent; }
    internal uint Id { get; }
    internal uint Kind { get; }
    internal int Offset { get; }
    internal int End { get; set; }
    internal uint? Parent { get; }
    internal bool IsShape => Kind is 0x01 or 0x20 or 0x30 or 0x31 or 0x10;
}

internal sealed class PublisherSourcePage {
    internal PublisherSourcePage(uint id, string name, uint? master, uint? background) { Id = id; Name = name; Master = master; Background = background; }
    internal uint Id { get; }
    internal string Name { get; }
    internal uint? Master { get; }
    internal uint? Background { get; }
    internal bool IsMaster => !string.IsNullOrEmpty(Name);
}

internal sealed class PublisherSourceShape {
    internal PublisherSourceShape(PublisherContentChunk chunk, IReadOnlyList<PublisherBlock> fields) { Chunk = chunk; Fields = fields; }
    internal PublisherContentChunk Chunk { get; }
    internal IReadOnlyList<PublisherBlock> Fields { get; }
    internal PublisherSourceTable? Table { get; set; }
    internal IReadOnlyList<uint> WrapObjects { get; set; } = Array.Empty<uint>();
    internal uint? Value(byte id) => PublisherContentsReader.Field(Fields, id)?.Value;
}
