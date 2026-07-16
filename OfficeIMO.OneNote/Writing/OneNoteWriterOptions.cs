namespace OfficeIMO.OneNote;

/// <summary>Controls native offline OneNote serialization.</summary>
public sealed class OneNoteWriterOptions {
    /// <summary>Maximum serialized output size. The default is 512 MiB.</summary>
    public long MaxOutputBytes { get; set; } = OneNoteReaderOptions.DefaultMaxInputBytes;

    /// <summary>Reads the generated artifact back before returning it.</summary>
    public bool ValidateRoundTrip { get; set; } = true;

    /// <summary>Maximum number of files emitted into a notebook directory or <c>.onepkg</c> archive.</summary>
    public int MaxPackageEntries { get; set; } = 10000;

    /// <summary>
    /// Preserves source objects, properties, and relationships that are not replaced by typed edits.
    /// This is enabled by default for sections loaded by OfficeIMO.
    /// </summary>
    public bool PreserveUnknownData { get; set; } = true;
}
