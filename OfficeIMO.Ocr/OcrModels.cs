using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Ocr;

/// <summary>Capabilities exposed by one configured OCR engine.</summary>
public sealed class OcrEngineCapabilities {
    /// <summary>Media types accepted by the engine. Empty means the engine does not pre-declare a restriction.</summary>
    public IReadOnlyList<string> SupportedMediaTypes { get; set; } = Array.Empty<string>();

    /// <summary>Language identifiers known to the engine. Empty means language discovery is provider-specific.</summary>
    public IReadOnlyList<string> SupportedLanguages { get; set; } = Array.Empty<string>();

    /// <summary>Whether line-level text spans can be returned.</summary>
    public bool SupportsLineSpans { get; set; }

    /// <summary>Whether word-level text spans can be returned.</summary>
    public bool SupportsWordSpans { get; set; }

    /// <summary>Whether character-level text spans can be returned.</summary>
    public bool SupportsCharacterSpans { get; set; }

    /// <summary>Whether confidence values can be returned.</summary>
    public bool SupportsConfidence { get; set; }

    /// <summary>Whether the same engine instance accepts concurrent recognition requests.</summary>
    public bool SupportsConcurrentRequests { get; set; }

    /// <summary>Creates an independent capability snapshot.</summary>
    public OcrEngineCapabilities Clone() => new OcrEngineCapabilities {
        SupportedMediaTypes = (SupportedMediaTypes ?? Array.Empty<string>()).ToArray(),
        SupportedLanguages = (SupportedLanguages ?? Array.Empty<string>()).ToArray(),
        SupportsLineSpans = SupportsLineSpans,
        SupportsWordSpans = SupportsWordSpans,
        SupportsCharacterSpans = SupportsCharacterSpans,
        SupportsConfidence = SupportsConfidence,
        SupportsConcurrentRequests = SupportsConcurrentRequests
    };
}

/// <summary>
/// One engine-neutral recognition request. Document formats map their source objects into this contract without
/// making OCR providers depend on Reader, PDF, Word, Excel, PowerPoint, or another format package.
/// </summary>
public sealed class OcrRequest {
    /// <summary>Validated raster payload supplied to the engine.</summary>
    public byte[] Payload { get; set; } = Array.Empty<byte>();

    /// <summary>IANA media type for <see cref="Payload"/>.</summary>
    public string MediaType { get; set; } = string.Empty;

    /// <summary>Original or synthetic file name, used only as provider input metadata.</summary>
    public string? FileName { get; set; }

    /// <summary>Stable caller-owned identifier for the source document or artifact.</summary>
    public string? SourceId { get; set; }

    /// <summary>Human-readable source path or logical name.</summary>
    public string? SourceName { get; set; }

    /// <summary>Stable caller-owned identifier for this recognition candidate.</summary>
    public string? CandidateId { get; set; }

    /// <summary>Caller-defined candidate kind such as image, page, slide, or worksheet-image.</summary>
    public string? CandidateKind { get; set; }

    /// <summary>One-based page or frame number within the source payload, when applicable.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Raster width in pixels, when known.</summary>
    public int? PixelWidth { get; set; }

    /// <summary>Raster height in pixels, when known.</summary>
    public int? PixelHeight { get; set; }

    /// <summary>Source region represented by the payload, when the caller extracted only part of an artifact.</summary>
    public OcrRegion? Region { get; set; }

    /// <summary>Coordinate unit used by <see cref="Region"/>.</summary>
    public OcrCoordinateUnit RegionCoordinateUnit { get; set; } = OcrCoordinateUnit.Pixels;

    /// <summary>Requested language tag or provider-specific language expression, when configured.</summary>
    public string? Language { get; set; }

    /// <summary>Provider-specific scalar options supplied by the host.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } =
        new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Recognition output returned by an OCR engine.</summary>
public sealed class OcrResult {
    /// <summary>Recognized plain text in source reading order.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Overall normalized confidence from zero through one, when available.</summary>
    public double? Confidence { get; set; }

    /// <summary>Detected or requested language identifier, when available.</summary>
    public string? Language { get; set; }

    /// <summary>Provider identifier reported by the engine.</summary>
    public string? Provider { get; set; }

    /// <summary>Provider model, engine, or trained-data identifier, when available.</summary>
    public string? Model { get; set; }

    /// <summary>Optional line, word, and character spans in provider reading order.</summary>
    public IReadOnlyList<OcrTextSpan> Spans { get; set; } = Array.Empty<OcrTextSpan>();

    /// <summary>Structured provider diagnostics produced during recognition.</summary>
    public IReadOnlyList<OcrDiagnostic> Diagnostics { get; set; } = Array.Empty<OcrDiagnostic>();
}

/// <summary>Granularity of one recognized OCR text span.</summary>
public enum OcrTextSpanLevel {
    /// <summary>One recognized text line.</summary>
    Line = 1,
    /// <summary>One recognized word or token.</summary>
    Word = 2,
    /// <summary>One recognized character or grapheme.</summary>
    Character = 3
}

/// <summary>Coordinate unit used by an OCR region.</summary>
public enum OcrCoordinateUnit {
    /// <summary>Source image pixels.</summary>
    Pixels = 0,
    /// <summary>Document points, where 72 points equal one inch.</summary>
    Points = 1,
    /// <summary>Normalized coordinates from zero through one.</summary>
    Normalized = 2
}

/// <summary>Axis-aligned region attached to an OCR request or recognized span.</summary>
public sealed class OcrRegion {
    /// <summary>Horizontal origin.</summary>
    public double X { get; set; }
    /// <summary>Vertical origin.</summary>
    public double Y { get; set; }
    /// <summary>Region width.</summary>
    public double Width { get; set; }
    /// <summary>Region height.</summary>
    public double Height { get; set; }
}

/// <summary>Detailed recognized line, word, or character with optional confidence and geometry.</summary>
public sealed class OcrTextSpan {
    /// <summary>Zero-based sequence within the provider's emitted reading order.</summary>
    public int Sequence { get; set; }
    /// <summary>Span granularity.</summary>
    public OcrTextSpanLevel Level { get; set; }
    /// <summary>Recognized text for this span.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Normalized confidence from zero through one, when available.</summary>
    public double? Confidence { get; set; }
    /// <summary>Detected or requested language identifier, when available.</summary>
    public string? Language { get; set; }
    /// <summary>One-based source page within a multi-page payload, when applicable.</summary>
    public int? PageNumber { get; set; }
    /// <summary>Provider-stable block identifier, when exposed.</summary>
    public string? BlockId { get; set; }
    /// <summary>Provider-stable paragraph identifier, when exposed.</summary>
    public string? ParagraphId { get; set; }
    /// <summary>Provider-stable line identifier, when exposed.</summary>
    public string? LineId { get; set; }
    /// <summary>Bounding region in <see cref="CoordinateUnit"/>, when available.</summary>
    public OcrRegion? Region { get; set; }
    /// <summary>Coordinate unit used by <see cref="Region"/>.</summary>
    public OcrCoordinateUnit CoordinateUnit { get; set; } = OcrCoordinateUnit.Pixels;
}

/// <summary>Severity of one provider diagnostic.</summary>
public enum OcrDiagnosticSeverity {
    /// <summary>Informational provider evidence.</summary>
    Info = 0,
    /// <summary>Recoverable condition that may reduce recognition quality.</summary>
    Warning = 1,
    /// <summary>Recognition error or invalid provider output.</summary>
    Error = 2
}

/// <summary>Engine-neutral provider diagnostic.</summary>
public sealed class OcrDiagnostic {
    /// <summary>Diagnostic severity.</summary>
    public OcrDiagnosticSeverity Severity { get; set; }
    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; set; } = string.Empty;
    /// <summary>Provider or component that emitted the diagnostic.</summary>
    public string? Source { get; set; }
    /// <summary>Whether recognition may continue after the condition.</summary>
    public bool IsRecoverable { get; set; } = true;
    /// <summary>Bounded scalar diagnostic attributes.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } =
        new Dictionary<string, string>(StringComparer.Ordinal);
}
