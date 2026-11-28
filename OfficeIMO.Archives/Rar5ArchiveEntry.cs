namespace OfficeIMO.Archives;

/// <summary>Metadata from the primary RAR5 file headers.</summary>
public sealed class Rar5ArchiveEntry {
    internal Rar5ArchiveEntry(int ordinal, string name, string path, ArchiveEntryKind kind, long length, long packedLength,
        int method, long dictionarySize, Rar5FileTimes times, Rar5UnixOwner? owner, Rar5Redirection? redirection, bool canRead) {
        Ordinal = ordinal;
        Name = name;
        RelativePath = path;
        Kind = kind;
        Length = length;
        PackedLength = packedLength;
        CompressionMethod = method;
        DictionarySize = dictionarySize;
        Times = times;
        UnixOwner = owner;
        Redirection = redirection;
        CanRead = canRead;
    }

    /// <summary>Zero-based position of the logical entry in the primary file headers.</summary>
    public int Ordinal { get; }
    /// <summary>Exact UTF-8 member name as decoded from the archive.</summary>
    public string Name { get; }
    /// <summary>Validated portable member path.</summary>
    public string RelativePath { get; }
    /// <summary>File, directory, link or special member kind.</summary>
    public ArchiveEntryKind Kind { get; }
    /// <summary>Uncompressed member size declared in the header.</summary>
    public long Length { get; }
    /// <summary>Total stored member data size over all volumes.</summary>
    public long PackedLength { get; }
    /// <summary>RAR5 method number. Zero indicates a stored member.</summary>
    public int CompressionMethod { get; }
    /// <summary>Declared dictionary size in bytes, or zero for a stored entry.</summary>
    public long DictionarySize { get; }
    /// <summary>Modification, creation and access times carried by the entry.</summary>
    public Rar5FileTimes Times { get; }
    /// <summary>UTC modification time, if present, rounded down to the nearest tick.</summary>
    public DateTimeOffset? LastWriteTime => Times.Modified?.UtcTime;
    /// <summary>Owner and group as recorded in the archive. Extraction does not apply ownership.</summary>
    public Rar5UnixOwner? UnixOwner { get; }
    /// <summary>Authored archive redirection. Extraction never creates or follows filesystem links.</summary>
    public Rar5Redirection? Redirection { get; }
    /// <summary>Whether the declared member profile can return file bytes. Listing alone does not check those bytes.</summary>
    public bool CanRead { get; }
}

/// <summary>The kinds of redirection in the RAR5 extra area.</summary>
public enum Rar5RedirectionKind {
    /// <summary>A Unix symbolic link.</summary>
    UnixSymbolicLink = 1,
    /// <summary>A Windows symbolic link.</summary>
    WindowsSymbolicLink = 2,
    /// <summary>A Windows directory junction.</summary>
    WindowsJunction = 3,
    /// <summary>A hard link.</summary>
    HardLink = 4,
    /// <summary>A file whose data are stored under another name in the same archive.</summary>
    FileCopy = 5
}

/// <summary>An inert filesystem redirection from a RAR5 entry.</summary>
public sealed class Rar5Redirection {
    internal Rar5Redirection(Rar5RedirectionKind kind, string target, bool targetIsDirectory) {
        Kind = kind;
        Target = target;
        TargetIsDirectory = targetIsDirectory;
    }
    /// <summary>The authored redirection kind.</summary>
    public Rar5RedirectionKind Kind { get; }
    /// <summary>Raw archive target string. It is never used as an operating-system path.</summary>
    public string Target { get; }
    /// <summary>Whether the record describes a directory target.</summary>
    public bool TargetIsDirectory { get; }
}

/// <summary>Optional RAR5 Unix ownership fields.</summary>
public sealed class Rar5UnixOwner {
    internal Rar5UnixOwner(byte[]? userName, byte[]? groupName, ulong? userId, ulong? groupId) {
        UserName = userName == null ? null : Array.AsReadOnly(userName);
        GroupName = groupName == null ? null : Array.AsReadOnly(groupName);
        UserId = userId;
        GroupId = groupId;
    }
    /// <summary>Authored user name bytes in the originating system's encoding.</summary>
    public IReadOnlyList<byte>? UserName { get; }
    /// <summary>Authored group name bytes in the originating system's encoding.</summary>
    public IReadOnlyList<byte>? GroupName { get; }
    /// <summary>Authored Unix user ID.</summary>
    public ulong? UserId { get; }
    /// <summary>Authored Unix group ID.</summary>
    public ulong? GroupId { get; }
}

/// <summary>A RAR5 timestamp with its original subsecond precision.</summary>
public readonly struct Rar5Timestamp {
    internal Rar5Timestamp(DateTimeOffset utcTime, int nanoseconds) {
        UtcTime = utcTime;
        Nanoseconds = nanoseconds;
    }
    /// <summary>UTC time rounded down to the nearest .NET tick.</summary>
    public DateTimeOffset UtcTime { get; }
    /// <summary>Nanoseconds within the second, including precision finer than a .NET tick.</summary>
    public int Nanoseconds { get; }
}

/// <summary>The RAR5 timestamps present on an entry.</summary>
public readonly struct Rar5FileTimes {
    internal Rar5FileTimes(Rar5Timestamp? modified, Rar5Timestamp? created, Rar5Timestamp? accessed) {
        Modified = modified;
        Created = created;
        Accessed = accessed;
    }
    /// <summary>Modification time.</summary>
    public Rar5Timestamp? Modified { get; }
    /// <summary>Creation time.</summary>
    public Rar5Timestamp? Created { get; }
    /// <summary>Access time.</summary>
    public Rar5Timestamp? Accessed { get; }
}

/// <summary>One RAR5 archive's checked file table and optional archive comment.</summary>
public sealed class Rar5ArchiveInfo {
    private readonly string? _comment;
    private readonly (int Method, long Dictionary)? _commentCompression;

    internal Rar5ArchiveInfo(IReadOnlyList<Rar5ArchiveEntry> entries, int volumeCount, string? comment,
        (int Method, long Dictionary)? commentCompression) {
        Entries = entries;
        VolumeCount = volumeCount;
        _comment = comment;
        _commentCompression = commentCompression;
    }
    /// <summary>Entries in primary header order, with split members joined.</summary>
    public IReadOnlyList<Rar5ArchiveEntry> Entries { get; }
    /// <summary>Number of supplied volumes.</summary>
    public int VolumeCount { get; }
    /// <summary>The CMT compression method. Null indicates no comment and zero indicates a stored comment.</summary>
    public int? CommentCompressionMethod => _commentCompression?.Method;
    /// <summary>The verified stored CMT service record as UTF-8 text, if present.</summary>
    /// <exception cref="Rar5CompressionNotSupportedException">The comment is compressed.</exception>
    public string? Comment => _commentCompression is { Method: > 0 } compression
        ? throw new Rar5CompressionNotSupportedException("CMT", compression.Method, compression.Dictionary)
        : _comment;
}
