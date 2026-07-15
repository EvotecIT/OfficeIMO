namespace OfficeIMO.Pdf;

/// <summary>Immutable, bounded projection of one PDF syntax value.</summary>
public sealed class PdfRawValue {
    internal PdfRawValue(
        PdfRawValueKind kind,
        double? number = null,
        bool? boolean = null,
        string? text = null,
        int? referenceObjectNumber = null,
        int? referenceGeneration = null,
        IReadOnlyList<PdfRawValue>? items = null,
        IReadOnlyDictionary<string, PdfRawValue>? entries = null,
        int? streamLength = null,
        bool streamDecodingFailed = false,
        bool isTruncated = false) {
        Kind = kind;
        Number = number;
        Boolean = boolean;
        Text = text;
        ReferenceObjectNumber = referenceObjectNumber;
        ReferenceGeneration = referenceGeneration;
        Items = items ?? Array.Empty<PdfRawValue>();
        Entries = entries ?? new System.Collections.ObjectModel.ReadOnlyDictionary<string, PdfRawValue>(new Dictionary<string, PdfRawValue>());
        StreamLength = streamLength;
        StreamDecodingFailed = streamDecodingFailed;
        IsTruncated = isTruncated;
    }

    /// <summary>Projected syntax kind.</summary>
    public PdfRawValueKind Kind { get; }
    /// <summary>Numeric value when <see cref="Kind"/> is <see cref="PdfRawValueKind.Number"/>.</summary>
    public double? Number { get; }
    /// <summary>Boolean value when <see cref="Kind"/> is <see cref="PdfRawValueKind.Boolean"/>.</summary>
    public bool? Boolean { get; }
    /// <summary>Bounded string or name value without PDF delimiters.</summary>
    public string? Text { get; }
    /// <summary>Referenced object number for indirect references.</summary>
    public int? ReferenceObjectNumber { get; }
    /// <summary>Referenced generation for indirect references.</summary>
    public int? ReferenceGeneration { get; }
    /// <summary>Bounded immutable array items.</summary>
    public IReadOnlyList<PdfRawValue> Items { get; }
    /// <summary>Bounded immutable dictionary entries.</summary>
    public IReadOnlyDictionary<string, PdfRawValue> Entries { get; }
    /// <summary>Parsed stream data length; stream bytes are deliberately not exposed.</summary>
    public int? StreamLength { get; }
    /// <summary>True when the parser retained undecoded stream bytes after a filter failure.</summary>
    public bool StreamDecodingFailed { get; }
    /// <summary>True when this projection omitted data because a configured bound was reached.</summary>
    public bool IsTruncated { get; }
}
