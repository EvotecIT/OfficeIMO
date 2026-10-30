using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Native XPS/OpenXPS semantic preservation and existing PDF output settings.</summary>
public sealed class XpsToPdfOptions {
    /// <summary>Maximum SVG characters projected across all PDF page occurrences. Default: 32 Mi characters.</summary>
    public long MaximumSvgCharacters { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum parsed path commands projected across all PDF page occurrences. Default: 200,000.</summary>
    public long MaximumPathCommands { get; set; } = 200_000;
    /// <summary>Maps authored native semantics strictly. False retains paint and markup-order Unicode without semantic reconstruction.</summary>
    public bool PreserveLogicalStructure { get; set; } = true;
    /// <summary>Existing PDF serialization, typography, compliance and encryption settings. Null uses the PDF engine defaults.</summary>
    public PdfCreateOptions? PdfOptions { get; set; }
    /// <summary>Returns an independent copy of the output settings.</summary>
    public XpsToPdfOptions Clone() => new() { PreserveLogicalStructure = PreserveLogicalStructure, PdfOptions = PdfOptions?.Clone(), MaximumSvgCharacters = MaximumSvgCharacters, MaximumPathCommands = MaximumPathCommands };
}
