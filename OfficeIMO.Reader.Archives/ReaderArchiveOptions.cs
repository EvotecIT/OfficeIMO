namespace OfficeIMO.Reader.Archives;

/// <summary>
/// Controls nested archives and optional container kinds for archive ingestion.
/// </summary>
public sealed class ReaderArchiveOptions {
    /// <summary>
    /// When true, archive members that are themselves archives (ZIP, 7z, CAB, TAR plain or compressed, single compressed
    /// files, and MSI when <see cref="ReadInstallerPackages"/> is set) are read recursively.
    /// </summary>
    public bool ReadNestedArchives { get; set; } = true;

    /// <summary>
    /// Maximum recursion depth for nested archives. 0 means top-level only.
    /// </summary>
    public int MaxNestedDepth { get; set; } = 2;

    /// <summary>
    /// Optional max byte size for materializing a nested archive.
    /// </summary>
    public long? MaxNestedArchiveBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>
    /// When true, Windows Installer packages (.msi) are read through their embedded cabinets, at the top level and
    /// when nested. Off by default: installers are large and rarely hold documents, so other packages are reported
    /// with a warning instead.
    /// </summary>
    public bool ReadInstallerPackages { get; set; }
}
