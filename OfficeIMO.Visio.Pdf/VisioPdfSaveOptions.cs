using OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Controls the loss-aware Visio-to-PDF projection.</summary>
public sealed class VisioPdfSaveOptions {
    /// <summary>Optional logical source name recorded in conversion diagnostics. Loaded documents use their associated path by default.</summary>
    public string? SourceName { get; set; }

    /// <summary>Visio-owned source projection settings.</summary>
    public VisioDocumentProjectionOptions? VisioOptions { get; set; }

    /// <summary>PDF-owned projection and generation settings.</summary>
    public PdfProjectionOptions? ProjectionOptions { get; set; }

    internal void Validate() {
        if (SourceName != null && string.IsNullOrWhiteSpace(SourceName)) {
            throw new ArgumentException("Source name is required for conversion diagnostics.", nameof(SourceName));
        }
    }
}
