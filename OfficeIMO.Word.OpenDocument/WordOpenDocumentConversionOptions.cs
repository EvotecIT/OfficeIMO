using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Controls optional content transferred by the Word/OpenDocument adapter.</summary>
public sealed class WordOpenDocumentConversionOptions {
    /// <summary>Controls whether reported conversion loss is returned or rejected.</summary>
    public OdfConversionLossPolicy LossPolicy { get; set; } = OdfConversionLossPolicy.ReportOnly;
    /// <summary>Copy embedded inline images when their bytes are available.</summary>
    public bool IncludeImages { get; set; } = true;
    /// <summary>Copy default headers and footers.</summary>
    public bool IncludeHeadersAndFooters { get; set; } = true;
}
