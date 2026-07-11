using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>Capabilities exposed by one configured OCR engine.</summary>
public sealed class OfficeOcrEngineCapabilities {
    /// <summary>Media types accepted by the engine. An empty collection means the provider does not pre-declare a restriction.</summary>
    public IReadOnlyList<string> SupportedMediaTypes { get; set; } = Array.Empty<string>();

    /// <summary>Language identifiers known to the engine. An empty collection means language discovery is provider-specific.</summary>
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
    public OfficeOcrEngineCapabilities Clone() {
        return new OfficeOcrEngineCapabilities {
            SupportedMediaTypes = (SupportedMediaTypes ?? Array.Empty<string>()).ToArray(),
            SupportedLanguages = (SupportedLanguages ?? Array.Empty<string>()).ToArray(),
            SupportsLineSpans = SupportsLineSpans,
            SupportsWordSpans = SupportsWordSpans,
            SupportsCharacterSpans = SupportsCharacterSpans,
            SupportsConfidence = SupportsConfidence,
            SupportsConcurrentRequests = SupportsConcurrentRequests
        };
    }
}

/// <summary>One bounded recognition request passed to an OCR engine.</summary>
public sealed class OfficeOcrEngineRequest {
    /// <summary>OCR candidate being resolved.</summary>
    public OfficeDocumentOcrCandidate Candidate { get; set; } = new OfficeDocumentOcrCandidate();

    /// <summary>Materialized asset associated with the candidate.</summary>
    public OfficeDocumentAsset Asset { get; set; } = new OfficeDocumentAsset();

    /// <summary>Validated asset payload supplied to the engine.</summary>
    public byte[] Payload { get; set; } = Array.Empty<byte>();

    /// <summary>Requested language or language expression, when configured.</summary>
    public string? Language { get; set; }

    /// <summary>Source document metadata for provider-side correlation.</summary>
    public OfficeDocumentSource Source { get; set; } = new OfficeDocumentSource();

    /// <summary>Provider-specific scalar options supplied by the host.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Recognition output returned by an OCR engine.</summary>
public sealed class OfficeOcrEngineResult {
    /// <summary>Recognized plain text in source reading order.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Overall normalized confidence from 0 through 1, when available.</summary>
    public double? Confidence { get; set; }

    /// <summary>Detected or requested language identifier, when available.</summary>
    public string? Language { get; set; }

    /// <summary>Provider identifier reported by the engine.</summary>
    public string? Provider { get; set; }

    /// <summary>Provider model, engine, or trained-data identifier, when available.</summary>
    public string? Model { get; set; }

    /// <summary>Optional line, word, and character spans in reading order.</summary>
    public IReadOnlyList<OfficeOcrTextSpan> Spans { get; set; } = Array.Empty<OfficeOcrTextSpan>();

    /// <summary>Structured provider diagnostics produced during recognition.</summary>
    public IReadOnlyList<OfficeDocumentDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentDiagnostic>();
}

/// <summary>Granularity of one recognized OCR text span.</summary>
public enum OfficeOcrTextSpanLevel {
    /// <summary>One recognized text line.</summary>
    Line = 1,
    /// <summary>One recognized word or token.</summary>
    Word = 2,
    /// <summary>One recognized character or grapheme.</summary>
    Character = 3
}

/// <summary>Coordinate unit used by an OCR span region.</summary>
public enum OfficeOcrCoordinateUnit {
    /// <summary>Source image pixels.</summary>
    Pixels = 0,
    /// <summary>Document points, where 72 points equal one inch.</summary>
    Points = 1,
    /// <summary>Normalized coordinates from 0 through 1.</summary>
    Normalized = 2
}

/// <summary>Detailed recognized line, word, or character with optional confidence and geometry.</summary>
public sealed class OfficeOcrTextSpan {
    /// <summary>Zero-based sequence within the provider's emitted reading order.</summary>
    public int Sequence { get; set; }

    /// <summary>Span granularity.</summary>
    public OfficeOcrTextSpanLevel Level { get; set; }

    /// <summary>Recognized text for this span.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Normalized confidence from 0 through 1, when available.</summary>
    public double? Confidence { get; set; }

    /// <summary>Detected or requested language identifier, when available.</summary>
    public string? Language { get; set; }

    /// <summary>One-based source page within a multi-page payload, when applicable.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Bounding region in <see cref="CoordinateUnit"/> when available.</summary>
    public OfficeDocumentRegion? Region { get; set; }

    /// <summary>Coordinate unit used by <see cref="Region"/>.</summary>
    public OfficeOcrCoordinateUnit CoordinateUnit { get; set; } = OfficeOcrCoordinateUnit.Pixels;
}

/// <summary>One candidate recognition paired with its engine output.</summary>
public sealed class OfficeDocumentOcrRecognition {
    /// <summary>Candidate identifier.</summary>
    public string CandidateId { get; set; } = string.Empty;

    /// <summary>Resolved asset identifier.</summary>
    public string AssetId { get; set; } = string.Empty;

    /// <summary>Engine output, including detailed spans and provider diagnostics.</summary>
    public OfficeOcrEngineResult Result { get; set; } = new OfficeOcrEngineResult();
}
