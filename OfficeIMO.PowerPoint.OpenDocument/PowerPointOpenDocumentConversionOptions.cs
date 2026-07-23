namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Controls optional content transferred by the PowerPoint/OpenDocument adapter.</summary>
public sealed class PowerPointOpenDocumentConversionOptions {
    /// <summary>Copy embedded images whose formats are supported by the target presentation model.</summary>
    public bool IncludeImages { get; set; } = true;
    /// <summary>Copy plain speaker-note text.</summary>
    public bool IncludeSpeakerNotes { get; set; } = true;
    /// <summary>Copy the common solid fill, outline, and text-run formatting subset.</summary>
    public bool IncludeBasicFormatting { get; set; } = true;

    /// <summary>Maximum rows allowed in a converted presentation table. Default: 4,096.</summary>
    public int MaxTableRows { get; set; } = 4_096;

    /// <summary>Maximum columns allowed in a converted presentation table. Default: 256.</summary>
    public int MaxTableColumns { get; set; } = 256;
}
