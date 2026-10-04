using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

/// <summary>Controls optional content transferred by the Word/OpenDocument adapter.</summary>
public sealed class WordOpenDocumentConversionOptions {
    /// <summary>Controls whether reported conversion loss is returned or rejected.</summary>
    public OdfConversionLossPolicy LossPolicy { get; set; } = OdfConversionLossPolicy.ReportOnly;
    /// <summary>Copy embedded inline images when their bytes are available.</summary>
    public bool IncludeImages { get; set; } = true;
    /// <summary>Maximum aggregate image bytes copied into one Word document; further images are reported as skipped.</summary>
    public long MaxConvertedImageBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Copy default headers and footers.</summary>
    public bool IncludeHeadersAndFooters { get; set; } = true;
    /// <summary>Maximum logical rows in one ODT table, checked before conversion.</summary>
    public int MaxTableRows { get; set; } = 4096;
    /// <summary>Maximum logical columns in one ODT table.</summary>
    public int MaxTableColumns { get; set; } = 256;
    /// <summary>Maximum cells allocated across all ODT tables, including repeated nested tables.</summary>
    public long MaxConvertedTableCells { get; set; } = 100_000;
    /// <summary>Maximum decoded table text characters after expanding repeated rows and cells, including nested tables.</summary>
    public long MaxConvertedTableTextCharacters { get; set; } = 16_000_000;
    /// <summary>Maximum nesting depth of ODT tables.</summary>
    public int MaxTableDepth { get; set; } = 32;
}
