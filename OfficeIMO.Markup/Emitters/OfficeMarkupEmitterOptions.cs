namespace OfficeIMO.Markup;

/// <summary>
/// Options shared by target-specific code emitters.
/// </summary>
public sealed class OfficeMarkupEmitterOptions {
    public string FilePathVariable { get; set; } = "filePath";
    public bool IncludeHeader { get; set; } = true;
}
