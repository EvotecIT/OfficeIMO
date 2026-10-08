using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Native XPS/OpenXPS semantic preservation and existing PDF output settings.</summary>
public sealed class XpsToPdfOptions {
    /// <summary>Maps authored native semantics strictly. False retains paint and markup-order Unicode without semantic reconstruction.</summary>
    public bool PreserveLogicalStructure { get; set; } = true;
    /// <summary>Existing PDF serialization, typography, compliance and encryption settings. Null uses the PDF engine defaults.</summary>
    public PdfOptions? PdfOptions { get; set; }
    /// <summary>Returns an independent copy of the output settings.</summary>
    public XpsToPdfOptions Clone() => new() { PreserveLogicalStructure = PreserveLogicalStructure, PdfOptions = PdfOptions?.Clone() };
}
