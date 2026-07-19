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

    internal ReaderOpenDocumentOptions Clone() => new ReaderOpenDocumentOptions {
        SheetName = SheetName,
        A1Range = A1Range,
        HeadersInFirstRow = HeadersInFirstRow,
        IncludeSpeakerNotes = IncludeSpeakerNotes,
        MaxXmlCharacters = MaxXmlCharacters
    };
}
