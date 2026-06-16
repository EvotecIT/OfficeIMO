namespace OfficeIMO.Rtf;

/// <summary>
/// File-system source flags used by RTF file table entries.
/// </summary>
[Flags]
public enum RtfFileSource {
    /// <summary>No file-system source flag is specified.</summary>
    None = 0,

    /// <summary>Macintosh file system.</summary>
    Mac = 1,

    /// <summary>MS-DOS file system.</summary>
    Dos = 2,

    /// <summary>NTFS file system.</summary>
    Ntfs = 4,

    /// <summary>HPFS file system.</summary>
    Hpfs = 8,

    /// <summary>Network file system source flag.</summary>
    Network = 16
}
