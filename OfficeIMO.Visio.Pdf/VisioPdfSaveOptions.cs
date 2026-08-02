using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Controls the loss-aware Visio-to-PDF projection.</summary>
public sealed class VisioPdfSaveOptions {
    /// <summary>Optional logical source name recorded in conversion diagnostics. Loaded documents use their associated path by default.</summary>
    public string? SourceName { get; set; }

    /// <summary>Shared limits and extraction settings used while normalizing the diagram.</summary>
    public ReaderOptions? ReaderOptions { get; set; }

    /// <summary>Visio-specific preview and extraction settings.</summary>
    public ReaderVisioOptions? VisioOptions { get; set; }

    /// <summary>Shared PDF projection and generation settings.</summary>
    public ReaderPdfProjectionOptions? ProjectionOptions { get; set; }

    internal void Validate() {
        if (SourceName != null && string.IsNullOrWhiteSpace(SourceName)) {
            throw new ArgumentException("Source name is required for conversion diagnostics.", nameof(SourceName));
        }
    }
}
