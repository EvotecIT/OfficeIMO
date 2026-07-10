namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Controls optional content transferred by the PowerPoint/OpenDocument adapter.</summary>
public sealed class PowerPointOpenDocumentConversionOptions {
    /// <summary>Copy embedded images whose formats are supported by the target presentation model.</summary>
    public bool IncludeImages { get; set; } = true;
    /// <summary>Copy plain speaker-note text.</summary>
    public bool IncludeSpeakerNotes { get; set; } = true;
    /// <summary>Copy the common solid fill, outline, and text-run formatting subset.</summary>
    public bool IncludeBasicFormatting { get; set; } = true;
}
