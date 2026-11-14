namespace OfficeIMO.Archives;

/// <summary>One file or directory to write into a ZIP or TAR archive.</summary>
public sealed class ArchiveWriteEntry {
    private readonly byte[]? _content;
    private readonly Func<Stream>? _open;

    private ArchiveWriteEntry(string path, bool isDirectory, byte[]? content, Func<Stream>? open, DateTimeOffset? lastWriteTime) {
        ArgumentNullException.ThrowIfNull(path);
        Path = path; IsDirectory = isDirectory; _content = content; _open = open; LastWriteTime = lastWriteTime;
    }

    /// <summary>A file whose complete content is already in memory.</summary>
    public static ArchiveWriteEntry File(string path, byte[] content, DateTimeOffset? lastWriteTime = null) {
        ArgumentNullException.ThrowIfNull(content);
        return new ArchiveWriteEntry(path, false, content, null, lastWriteTime);
    }

    /// <summary>A file read from a stream that the writer opens when it reaches the entry and disposes afterwards.</summary>
    public static ArchiveWriteEntry File(string path, Func<Stream> openContent, DateTimeOffset? lastWriteTime = null) {
        ArgumentNullException.ThrowIfNull(openContent);
        return new ArchiveWriteEntry(path, false, null, openContent, lastWriteTime);
    }

    /// <summary>An explicit directory record.</summary>
    public static ArchiveWriteEntry Directory(string path, DateTimeOffset? lastWriteTime = null) =>
        new(path, true, null, null, lastWriteTime);

    /// <summary>Relative member path using '/' (or '\') separators. It must pass the same safety checks as extraction.</summary>
    public string Path { get; }
    /// <summary>Whether the entry is a directory record.</summary>
    public bool IsDirectory { get; }
    /// <summary>Entry modification time. Ignored when the writer options fix a timestamp.</summary>
    public DateTimeOffset? LastWriteTime { get; }

    internal Stream OpenContent() => _content != null ? new MemoryStream(_content, writable: false)
        : _open!() ?? throw new InvalidOperationException("The entry content factory returned null: " + Path);

    internal long? KnownLength => _content?.LongLength;
}
