namespace OfficeIMO.DjVu;

internal sealed class DjVuChunk {
    internal readonly byte[] Source;
    internal readonly string Id;
    internal readonly string? FormType;
    internal readonly int Offset;
    internal readonly int Length;
    internal readonly List<DjVuChunk> Children;
    internal int HeaderOffset => Offset - 8;
    internal DjVuChunk(byte[] source, string id, string? formType, int offset, int length, List<DjVuChunk> children) {
        Source = source; Id = id; FormType = formType; Offset = offset; Length = length; Children = children;
    }
}

internal sealed class DjVuComponent {
    internal readonly string Id;
    internal readonly string Name;
    internal readonly string Title;
    internal readonly int Kind;
    internal readonly DjVuChunk Form;
    internal DjVuComponent(string id, string name, string title, int kind, DjVuChunk form) {
        Id = id; Name = name; Title = title; Kind = kind; Form = form;
    }
}

internal static class DjVuBinary {
    internal static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
    internal static int U16(byte[] bytes, int offset) => bytes[offset] << 8 | bytes[offset + 1];
    internal static int U24(byte[] bytes, int offset) => bytes[offset] << 16 | bytes[offset + 1] << 8 | bytes[offset + 2];
    internal static uint U32(byte[] bytes, int offset) => (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16 | (uint)bytes[offset + 2] << 8 | bytes[offset + 3];
    internal static string Id(byte[] bytes, int offset) {
        for (int i = 0; i < 4; i++) if (bytes[offset + i] < 32 || bytes[offset + i] > 126) throw new InvalidDataException("Invalid IFF chunk identifier.");
        return Encoding.ASCII.GetString(bytes, offset, 4);
    }
}
