namespace OfficeIMO.Chm;

/// <summary>An immutable entry backed by the book's bounded, in-memory section snapshot.</summary>
public sealed class ChmEntry {
    private readonly byte[] _section;
    private readonly int _offset;
    internal ChmEntry(string name, string path, int section, byte[] data, int offset, int length) {
        Name = name; Path = path; Section = section; _section = data; _offset = offset; Length = length;
    }
    /// <summary>Exact UTF-8 name from the CHM directory.</summary>
    public string Name { get; }
    /// <summary>Canonical archive path, with a leading slash for content entries.</summary>
    public string Path { get; }
    /// <summary>Storage section: zero is uncompressed and one is MSCompressed.</summary>
    public int Section { get; }
    /// <summary>Expanded entry bytes.</summary>
    public int Length { get; }
    /// <summary>Whether this entry belongs to CHM control, search, or navigation metadata.</summary>
    public bool IsSystem => Path.StartsWith("::", StringComparison.Ordinal) || Path.StartsWith("/#", StringComparison.Ordinal) || Path.StartsWith("/$", StringComparison.Ordinal);
    /// <summary>Whether this entry is a directory marker.</summary>
    public bool IsDirectory => Path.EndsWith("/", StringComparison.Ordinal);
    /// <summary>Returns an independent byte copy. Changes to it cannot modify the book.</summary>
    public byte[] GetBytes() {
        var result = new byte[Length];
        Buffer.BlockCopy(_section, _offset, result, 0, Length);
        return result;
    }
    /// <summary>Opens a read-only stream whose lifetime is independent of the source stream.</summary>
    public Stream OpenRead() => new MemoryStream(_section, _offset, Length, false, false);
}
