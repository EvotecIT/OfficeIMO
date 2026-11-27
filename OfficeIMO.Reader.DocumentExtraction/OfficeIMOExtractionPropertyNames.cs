namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>
/// Keys that the OfficeIMO provider writes into <c>AIContent.AdditionalProperties</c> on element payloads and into
/// <c>ExtractionDocument.Metadata</c> for OfficeIMO data that the shared document model has no dedicated slot for.
/// </summary>
/// <remarks>
/// Values are plain strings, numbers, booleans or lists of string-keyed dictionaries, so they serialize without
/// OfficeIMO types. The full OfficeIMO model stays available through <c>RawRepresentation</c>.
/// </remarks>
public static class OfficeIMOExtractionPropertyNames {
    /// <summary>OfficeIMO block kind of a text element, such as <c>paragraph</c>, <c>list-item</c> or <c>speaker-notes</c>.</summary>
    public const string BlockKind = "officeimo.blockKind";

    /// <summary>Stable OfficeIMO block identifier within the read result.</summary>
    public const string BlockId = "officeimo.blockId";

    /// <summary>Stable source anchor of a block, table or asset, when the reader supplied one.</summary>
    public const string BlockAnchor = "officeimo.blockAnchor";

    /// <summary>Heading path of the element, for example <c>Intro &gt; Scope</c>.</summary>
    public const string HeadingPath = "officeimo.headingPath";

    /// <summary>Nesting depth of a list item (zero-based when the reader reports it that way).</summary>
    public const string ListLevel = "officeimo.listLevel";

    /// <summary>Displayed list marker, such as <c>1.</c> or a bullet character.</summary>
    public const string ListMarker = "officeimo.listMarker";

    /// <summary>Resolved numeric index of an ordered list item.</summary>
    public const string ListIndex = "officeimo.listIndex";

    /// <summary>Worksheet name for spreadsheet content.</summary>
    public const string Sheet = "officeimo.sheet";

    /// <summary>A1 cell range for spreadsheet content, for example <c>A1:C10</c>.</summary>
    public const string CellRange = "officeimo.a1Range";

    /// <summary>One-based slide number for presentation content.</summary>
    public const string Slide = "officeimo.slide";

    /// <summary>Table title supplied by the source format.</summary>
    public const string TableTitle = "officeimo.tableTitle";

    /// <summary>Table kind reported by the reader.</summary>
    public const string TableKind = "officeimo.tableKind";

    /// <summary>True when the reader truncated the table rows to its configured limit.</summary>
    public const string TableTruncated = "officeimo.tableTruncated";

    /// <summary>Row count of the source table before truncation.</summary>
    public const string TableTotalRowCount = "officeimo.tableTotalRowCount";

    /// <summary>Asset identifier of an image element.</summary>
    public const string AssetId = "officeimo.assetId";

    /// <summary>Asset title, kept separately from the alternative text.</summary>
    public const string AssetTitle = "officeimo.assetTitle";

    /// <summary>Asset file name inside the source package.</summary>
    public const string AssetFileName = "officeimo.assetFileName";

    /// <summary>Image width in pixels.</summary>
    public const string ImageWidth = "officeimo.imageWidth";

    /// <summary>Image height in pixels.</summary>
    public const string ImageHeight = "officeimo.imageHeight";

    /// <summary>OCR provider that produced the image text.</summary>
    public const string OcrProvider = "officeimo.ocrProvider";

    /// <summary>OCR confidence between 0 and 1.</summary>
    public const string OcrConfidence = "officeimo.ocrConfidence";

    /// <summary>
    /// Document metadata key holding hyperlinks and internal links as a list of dictionaries with <c>kind</c>, <c>text</c>,
    /// <c>uri</c>, <c>destinationName</c>, <c>destinationPage</c>, <c>page</c> and <c>blockAnchor</c> entries.
    /// </summary>
    public const string Links = "officeimo.links";

    /// <summary>
    /// Document metadata key holding form fields as a list of dictionaries with <c>name</c>, <c>kind</c>, <c>value</c>,
    /// <c>readOnly</c>, <c>required</c> and <c>page</c> entries.
    /// </summary>
    public const string FormFields = "officeimo.formFields";

    /// <summary>
    /// Document metadata key holding reader diagnostics as a list of dictionaries with <c>severity</c>, <c>category</c>,
    /// <c>code</c> and <c>message</c> entries.
    /// </summary>
    public const string Diagnostics = "officeimo.diagnostics";

    /// <summary>Document metadata key holding the detected OfficeIMO input kind, such as <c>Word</c> or <c>Pdf</c>.</summary>
    public const string InputKind = "officeimo.inputKind";

    /// <summary>Document metadata key holding the SHA-256 hash of the source bytes, when computed.</summary>
    public const string SourceHash = "officeimo.sourceHash";

    /// <summary>
    /// Key in <c>DocumentExtractionOptions.AdditionalProperties</c> that supplies per-call
    /// <see cref="OfficeIMO.Reader.ReaderOptions"/>. The value is borrowed for the duration of the call.
    /// </summary>
    public const string ReaderOptionsOverride = "officeimo.readerOptions";
}
