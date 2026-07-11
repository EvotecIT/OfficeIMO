namespace OfficeIMO.Word.OpenDocument;

/// <summary>Controls optional content transferred by the Word/OpenDocument adapter.</summary>
public sealed class WordOpenDocumentConversionOptions {
    /// <summary>Copy embedded inline images when their bytes are available.</summary>
    public bool IncludeImages { get; set; } = true;
    /// <summary>Copy default headers and footers.</summary>
    public bool IncludeHeadersAndFooters { get; set; } = true;
}
