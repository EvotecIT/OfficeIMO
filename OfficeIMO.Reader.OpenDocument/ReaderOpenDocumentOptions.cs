namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Controls format-specific OpenDocument text, spreadsheet, and presentation projection.</summary>
public sealed class ReaderOpenDocumentOptions {
    /// <summary>Optional ODS sheet name. Null selects every sheet.</summary>
    public string? SheetName { get; set; }

    /// <summary>Optional ODS A1 range. Null uses each selected sheet's used range.</summary>
    public string? A1Range { get; set; }

    /// <summary>Treats the first selected ODS row as column names.</summary>
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>Includes ODP speaker notes.</summary>
    public bool IncludeSpeakerNotes { get; set; } = true;

    /// <summary>Maximum decoded XML characters accepted by the OpenDocument parser.</summary>
    public long? MaxXmlCharacters { get; set; } = 10_000_000L;

    /// <summary>Maximum text characters extracted across the document before building joined text or Markdown. Repeated cells consume this budget for each logical occurrence.</summary>
    public long MaxExtractedCharacters { get; set; } = 16_000_000L;

    /// <summary>Maximum embedded image placements retained in a rich result. Default: 512.</summary>
    public int MaxImageAssets { get; set; } = 512;
    /// <summary>Maximum bytes copied from one embedded image. Default: 32 MiB.</summary>
    public long MaxImageBytes { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum unique image payload bytes retained. Repeated placements share their payload. Default: 128 MiB.</summary>
    public long MaxTotalImageBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum hyperlinks retained by native rich projection. Default: 4,096.</summary>
    public int MaxLinks { get; set; } = 4096;
    /// <summary>Maximum combined hyperlink target and display characters. Default: one million.</summary>
    public long MaxLinkCharacters { get; set; } = 1_000_000;
    /// <summary>Maximum native rich text blocks retained, including master and note stories. Default: 65,536.</summary>
    public int MaxRichBlocks { get; set; } = 65_536;

    internal ReaderOpenDocumentOptions Clone() => new ReaderOpenDocumentOptions {
        SheetName = SheetName,
        A1Range = A1Range,
        HeadersInFirstRow = HeadersInFirstRow,
        IncludeSpeakerNotes = IncludeSpeakerNotes,
        MaxImageAssets = MaxImageAssets,
        MaxImageBytes = MaxImageBytes,
        MaxTotalImageBytes = MaxTotalImageBytes,
        MaxLinks = MaxLinks,
        MaxLinkCharacters = MaxLinkCharacters,
        MaxRichBlocks = MaxRichBlocks,
        MaxXmlCharacters = MaxXmlCharacters,
        MaxExtractedCharacters = MaxExtractedCharacters
    };
}
