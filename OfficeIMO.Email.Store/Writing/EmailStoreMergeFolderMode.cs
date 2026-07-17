namespace OfficeIMO.Email.Store;

/// <summary>Controls how source folder hierarchies are projected into a merged PST.</summary>
public enum EmailStoreMergeFolderMode {
    /// <summary>Creates one isolated destination root per source and preserves its hierarchy.</summary>
    SeparateSourceRoots = 0,
    /// <summary>Merges case-insensitively equal folder paths from every source.</summary>
    MergeByFolderPath = 1,
    /// <summary>Places every source item directly in the destination root.</summary>
    Flatten = 2
}
