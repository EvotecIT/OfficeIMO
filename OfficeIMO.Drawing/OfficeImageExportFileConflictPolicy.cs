namespace OfficeIMO.Drawing;

/// <summary>Controls how image export handles an existing destination file.</summary>
public enum OfficeImageExportFileConflictPolicy {
    /// <summary>Fail without replacing the existing file.</summary>
    FailIfExists,
    /// <summary>Atomically replace the existing file.</summary>
    Replace,
    /// <summary>Choose a unique suffixed file name in the same directory.</summary>
    CreateUnique
}
