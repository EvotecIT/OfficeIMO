namespace OfficeIMO.Rtf;

/// <summary>
/// File reference from the RTF <c>{\*\filetbl ...}</c> destination.
/// </summary>
public sealed class RtfFileReference {
    /// <summary>
    /// Initializes a file-table reference.
    /// </summary>
    public RtfFileReference(int id, string path) {
        Id = id;
        Path = path ?? throw new ArgumentNullException(nameof(path));
    }

    /// <summary>RTF file identifier from <c>\fid</c>.</summary>
    public int Id { get; }

    /// <summary>Referenced path text.</summary>
    public string Path { get; set; }

    /// <summary>Character offset where the relative portion of the path starts, from <c>\frelative</c>.</summary>
    public int? RelativePathStart { get; set; }

    /// <summary>Operating-system-specific file number from <c>\fosnum</c>.</summary>
    public int? OperatingSystemNumber { get; set; }

    /// <summary>File-system source flags such as <c>\fvalidntfs</c> and <c>\fnetwork</c>.</summary>
    public RtfFileSource Sources { get; set; }
}
