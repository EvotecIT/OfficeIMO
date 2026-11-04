using OfficeIMO.ContentSafety;
using OfficeIMO.Provenance;
using OfficeIMO.Reader;
using OfficeIMO.Security;
using OfficeIMO.Word;
using OfficeIMO.Pdf;

namespace OfficeIMO;

/// <summary>Selected native content, provenance and signature reports for an unchanged source document.</summary>
public sealed class OfficeDocumentInfo {
    internal OfficeDocumentInfo(string source) { Source = source; }
    /// <summary>Path or caller-supplied name of the inspected source.</summary>
    public string Source { get; }
    /// <summary>Source path or logical filename; null for an unassociated live document.</summary>
    public string? Path => string.IsNullOrEmpty(Source) ? null : Source;
    private ReaderInputKind _format;
    /// <summary>Effective format identified by the native model or requested detection policy.</summary>
    public ReaderInputKind Format { get => _format != ReaderInputKind.Unknown ? _format : Content?.Kind ?? ReaderInputKind.Unknown; internal set => _format = value; }
    /// <summary>Format identified from content, independent of the filename.</summary>
    public ReaderInputKind DetectedFormat => Detection?.ContentKind ?? ReaderInputKind.Unknown;
    /// <summary>Confidence of source detection; live models do not run file detection.</summary>
    public ReaderDetectionConfidence Confidence => Detection?.Confidence ?? ReaderDetectionConfidence.None;
    /// <summary>Bounded source detection evidence, when requested.</summary>
    public ReaderDetectionResult? Detection { get; internal set; }
    /// <summary>Native page count when known. Word uses recorded statistics or explicit structure, without repagination.</summary>
    public int? PageCount { get; internal set; }
    /// <summary>Number of source worksheets, including empty and hidden sheets, when known.</summary>
    public int? SheetCount { get; internal set; }
    /// <summary>Number of source slides, including hidden slides, when known.</summary>
    public int? SlideCount { get; internal set; }
    /// <summary>Encryption evidence; null when the selected inspection cannot establish it.</summary>
    public bool? IsEncrypted { get; internal set; }
    /// <summary>Source filesystem modification time in UTC, when inspecting a path.</summary>
    public DateTime? Modified { get; internal set; }
    /// <summary>Detached snapshot produced by the native format's inspection API, when requested.</summary>
    public object? Details { get; internal set; }
    /// <summary>Native content and metadata, when requested.</summary>
    public OfficeDocumentReadResult? Content { get; internal set; }
    /// <summary>PDF read and rewrite capabilities, when requested.</summary>
    public PdfDocumentPreflight? PdfPreflight { get; internal set; }
    /// <summary>Concealed-content evidence, when requested.</summary>
    public OfficeContentSafetyReport? ContentSafety { get; internal set; }
    /// <summary>Source provenance evidence, when requested.</summary>
    public OfficeProvenanceReport? Provenance { get; internal set; }
    /// <summary>OPC package signature metadata, when requested.</summary>
    public OfficeIMO.Security.OfficePackageSignatureInfo? PackageSignatures { get; internal set; }
    /// <summary>VBA signature metadata, when requested.</summary>
    public OfficeVbaSignatureInfo? VbaSignatures { get; internal set; }
    /// <summary>Word macro project signature evidence, when requested.</summary>
    public WordMacroProjectSignatureInfo? WordMacroSignatures { get; internal set; }
}
