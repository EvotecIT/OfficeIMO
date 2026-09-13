namespace OfficeIMO.Project;

/// <summary>Project 98 field descriptor definitions, independent of producer record payloads and process pointers.</summary>
internal sealed class ProjectNativeLegacy8Schema {
    private readonly int _flags, _presence, _indexed;
    private readonly List<byte[]> _fields = new List<byte[]>();
    internal int Width { get; }
    internal ProjectNativeLegacy8Schema(int width, int flags, int presence, int indexed) {
        Width = width; _flags = flags; _presence = presence; _indexed = indexed;
    }
    internal void Field(uint id, ushort offset, int attributes, ushort type, ushort size, ushort bit) {
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(0); writer.Write(offset); writer.Write((ushort)0); writer.Write(id); writer.Write(attributes);
        writer.Write(type); writer.Write(size); writer.Write(bit); writer.Write((ushort)0); _fields.Add(stream.ToArray());
    }
    internal byte[] Serialize(long budget, CancellationToken token) {
        var properties = new ProjectNativePropertySet();
        properties.Set(1, 24, _fields.SelectMany(f => f).ToArray());
        properties.Set(2, 4, BitConverter.GetBytes((long)Width));
        properties.Set(6, 4, BitConverter.GetBytes(_flags)); properties.Set(7, 4, BitConverter.GetBytes(_presence));
        properties.Set(8, 4, BitConverter.GetBytes(_indexed)); properties.Set(3, 4, BitConverter.GetBytes(1));
        properties.Set(4, 4, BitConverter.GetBytes(_indexed < 0 ? 0 : 1)); properties.Set(5, 4, BitConverter.GetBytes(Width));
        return properties.Serialize(budget, token);
    }
}

internal static partial class ProjectNativeSchema8Catalog {
    internal static (string Name, ProjectNativeLegacy8Schema Schema)[] Tables() => new[] {
        ("Task", Task()), ("Rsc", Resource()), ("Cal", Calendar()), ("Assn", Assignment()), ("Cons", Dependency())
    };
}
