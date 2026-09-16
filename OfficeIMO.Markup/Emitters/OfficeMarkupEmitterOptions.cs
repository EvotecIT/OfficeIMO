namespace OfficeIMO.Markup;

/// <summary>
/// Options shared by target-specific code emitters.
/// </summary>
public sealed class OfficeMarkupEmitterOptions {
    /// <summary>Gets or sets the output-path variable text inserted into generated code; defaults to <c>filePath</c>.</summary>
    /// <remarks>The PowerShell emitter uses <c>$FilePath</c> when the entire options object is omitted.</remarks>
    public string FilePathVariable { get; set; } = "filePath";
    /// <summary>Gets or sets whether generated code begins with explanatory comments; defaults to true.</summary>
    public bool IncludeHeader { get; set; } = true;
}
