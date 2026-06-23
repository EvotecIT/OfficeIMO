namespace OfficeIMO.Pdf;

/// <summary>Summary of a PDF stream object discovered during diagnostics.</summary>
public sealed class PdfStreamDiagnostic {
    internal PdfStreamDiagnostic(
        int objectNumber,
        int generation,
        string kind,
        string? type,
        string? subtype,
        IReadOnlyList<string> filters,
        long length,
        bool decoded,
        bool decodingFailed,
        string? decodingError,
        int? width,
        int? height,
        int? bitsPerComponent,
        string hash) {
        ObjectNumber = objectNumber;
        Generation = generation;
        Kind = kind;
        Type = type;
        Subtype = subtype;
        Filters = filters;
        Length = length;
        Decoded = decoded;
        DecodingFailed = decodingFailed;
        DecodingError = decodingError;
        Width = width;
        Height = height;
        BitsPerComponent = bitsPerComponent;
        Hash = hash;
    }

    /// <summary>Object number containing the stream.</summary>
    public int ObjectNumber { get; }

    /// <summary>PDF object generation.</summary>
    public int Generation { get; }

    /// <summary>Friendly stream kind, usually the subtype, type, or Stream.</summary>
    public string Kind { get; }

    /// <summary>Dictionary /Type name, when present.</summary>
    public string? Type { get; }

    /// <summary>Dictionary /Subtype name, when present.</summary>
    public string? Subtype { get; }

    /// <summary>Dictionary /Filter names, when present.</summary>
    public IReadOnlyList<string> Filters { get; }

    /// <summary>Length of the stream bytes retained by the parser.</summary>
    public long Length { get; }

    /// <summary>True when the stream was decoded by the reader.</summary>
    public bool Decoded { get; }

    /// <summary>True when stream decoding failed and original bytes were retained.</summary>
    public bool DecodingFailed { get; }

    /// <summary>Stream decoding error, when available.</summary>
    public string? DecodingError { get; }

    /// <summary>Image width, when present.</summary>
    public int? Width { get; }

    /// <summary>Image height, when present.</summary>
    public int? Height { get; }

    /// <summary>Image bits per component, when present.</summary>
    public int? BitsPerComponent { get; }

    /// <summary>Stable non-cryptographic hash of the retained stream bytes.</summary>
    public string Hash { get; }
}
