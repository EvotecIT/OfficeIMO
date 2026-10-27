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
