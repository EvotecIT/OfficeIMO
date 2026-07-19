namespace OfficeIMO.Reader.PowerPoint;

/// <summary>Controls PowerPoint presentation projection into Reader chunks.</summary>
public sealed class ReaderPowerPointOptions {
    /// <summary>Includes speaker notes when present.</summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>Includes tables in the Markdown projection.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>Includes hidden shapes in the projection.</summary>
    public bool IncludeHiddenShapes { get; set; } = true;
}
