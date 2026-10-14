using OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Selects the PDF content contract.</summary>
public enum VisioPdfProjectionMode {
    /// <summary>Searchable semantic text and topology with optional preview assets.</summary>
    SemanticReport,
    /// <summary>One physical PDF page per source page, using cached diagram Drawing scenes.</summary>
    DiagramPages
}

/// <summary>Controls the loss-aware Visio-to-PDF projection.</summary>
public sealed class VisioToPdfOptions {
    /// <summary>PDF content contract. Defaults to the existing semantic report.</summary>
    public VisioPdfProjectionMode Mode { get; set; } = VisioPdfProjectionMode.SemanticReport;

    /// <summary>Optional logical source name recorded in conversion diagnostics. Loaded documents use their associated path by default.</summary>
    public string? SourceName { get; set; }

    /// <summary>Visio-owned source projection settings.</summary>
    public VisioDocumentProjectionOptions? VisioOptions { get; set; }

    /// <summary>PDF-owned projection and generation settings.</summary>
    public PdfProjectionOptions? ProjectionOptions { get; set; }

    /// <summary>Visio-owned diagram projection settings, used only by <see cref="VisioPdfProjectionMode.DiagramPages"/>.</summary>
    public VisioDrawingOptions? DrawingOptions { get; set; }

    /// <summary>PDF generation settings for diagram pages. Source page sizes and zero margins take precedence.</summary>
    public PdfOptions? PdfOptions { get; set; }

    internal void Validate() {
        if (!Enum.IsDefined(typeof(VisioPdfProjectionMode), Mode)) throw new ArgumentOutOfRangeException(nameof(Mode));
        if (SourceName != null && string.IsNullOrWhiteSpace(SourceName)) {
            throw new ArgumentException("Source name is required for conversion diagnostics.", nameof(SourceName));
        }
        if (Mode == VisioPdfProjectionMode.SemanticReport && (DrawingOptions != null || PdfOptions != null))
            throw new ArgumentException("Diagram settings require DiagramPages mode.");
        if (Mode == VisioPdfProjectionMode.DiagramPages && (VisioOptions != null || ProjectionOptions != null))
            throw new ArgumentException("Semantic projection settings require SemanticReport mode.");
    }
}
