namespace OfficeIMO.Project;

/// <summary>A generation-specific native table layout, independent of document data.</summary>
internal sealed class ProjectNativeSchema {
    private readonly ProjectNativeSchemaField[] _fields;
    internal ProjectNativeSchema(int primaryCount, ProjectNativeSchemaField[] fields) { PrimaryCount = primaryCount; _fields = fields; }
    internal int PrimaryCount { get; }
    internal int Count => _fields.Length;
    internal int Width => _fields.Take(PrimaryCount).Where(f => f.Source == 10).Select(f => f.Offset + f.Size).DefaultIfEmpty(0).Max();
    internal int SecondaryWidth => _fields.Skip(PrimaryCount).Where(f => f.Source == 10).Select(f => f.Offset + f.Size).DefaultIfEmpty(0).Max();
    internal byte[] Bytes(bool extended) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        foreach (var field in _fields.Take(extended ? _fields.Length : PrimaryCount)) {
            writer.Write(field.Mask); writer.Write(field.Offset); writer.Write(field.Source); writer.Write(field.Id);
            writer.Write(field.Attributes); writer.Write(field.Type); writer.Write(field.Size); writer.Write(field.Bit);
        }
        return stream.ToArray();
    }
}
internal readonly struct ProjectNativeSchemaField {
    internal readonly uint Id, Mask;
    internal readonly int Offset, Source, Attributes, Bit;
    internal readonly ushort Type, Size;
    internal ProjectNativeSchemaField(uint id, uint mask, int offset, int source, int attributes, ushort type, ushort size, int bit) {
        Id = id; Mask = mask; Offset = offset; Source = source; Attributes = attributes; Type = type; Size = size; Bit = bit;
    }
}
