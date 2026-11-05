namespace OfficeIMO.OneNote;

public sealed partial class OneNoteSection {
    /// <summary>Gets the associated filesystem path, separate from package-entry provenance in SourcePath.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Creates an empty native section in memory.</summary>
    public static OneNoteSection Create() => new OneNoteSection();

    /// <summary>Creates an empty section associated with a path without writing it.</summary>
    public static OneNoteSection Create(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A section path is required.", nameof(path));
        return new OneNoteSection { FilePath = Path.GetFullPath(path) };
    }

    /// <summary>Loads a native section and associates its filesystem path for subsequent saves.</summary>
    public static OneNoteSection Load(string path, OneNoteLoadOptions? options = null) {
        var section = OneNoteSectionReader.Read(path, options);
        section.FilePath = Path.GetFullPath(path);
        return section;
    }

    /// <summary>Loads a native section from a caller-owned seekable stream, without a path association.</summary>
    public static OneNoteSection Load(Stream stream, OneNoteLoadOptions? options = null) => OneNoteSectionReader.Read(stream, options);

    /// <summary>Saves to the last successful path load or save.</summary>
    public void Save(OneNoteSaveOptions? options = null) =>
        Save(FilePath ?? throw new InvalidOperationException("The section has no associated path. Save to a path first."), options);

    /// <summary>Writes a native section at the caller stream's current position and leaves it open.</summary>
    public void Save(Stream destination, OneNoteSaveOptions? options = null) => OneNoteSectionWriter.Write(this, destination, options);

    /// <summary>Completes this managed section's lifetime; no source file or caller stream is retained.</summary>
    public void Dispose() { }
}
