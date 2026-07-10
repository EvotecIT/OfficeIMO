using System;

namespace OfficeIMO.Reader;

/// <summary>Options for deterministic block, heading, and list normalization.</summary>
public sealed class OfficeDocumentBlockNormalizationOptions {
    /// <summary>Trim block text.</summary>
    public bool TrimText { get; set; } = true;

    /// <summary>Trim and lowercase block kind identifiers.</summary>
    public bool NormalizeKinds { get; set; } = true;

    /// <summary>Trim list and leader markers.</summary>
    public bool TrimMarkers { get; set; } = true;

    /// <summary>Clamp heading levels to 1-6 and list levels to a minimum of 1.</summary>
    public bool NormalizeLevels { get; set; } = true;

    internal OfficeDocumentBlockNormalizationOptions Clone() => new OfficeDocumentBlockNormalizationOptions {
        TrimText = TrimText,
        NormalizeKinds = NormalizeKinds,
        TrimMarkers = TrimMarkers,
        NormalizeLevels = NormalizeLevels
    };
}

/// <summary>Normalizes shared logical blocks without invoking a format-specific parser.</summary>
public sealed class OfficeDocumentBlockNormalizationProcessor : OfficeDocumentProcessorBase {
    private readonly OfficeDocumentBlockNormalizationOptions _options;

    /// <summary>Creates the processor.</summary>
    public OfficeDocumentBlockNormalizationProcessor(OfficeDocumentBlockNormalizationOptions? options = null)
        : base("officeimo.reader.normalize-blocks") {
        _options = (options ?? new OfficeDocumentBlockNormalizationOptions()).Clone();
    }

    /// <inheritdoc />
    public override OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        foreach (OfficeDocumentBlock block in OfficeDocumentModelTraversal.Blocks(document)) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (_options.TrimText) block.Text = (block.Text ?? string.Empty).Trim();
            if (_options.NormalizeKinds) block.Kind = (block.Kind ?? string.Empty).Trim().ToLowerInvariant();
            if (_options.TrimMarkers && block.Marker != null) block.Marker = block.Marker.Trim();
            if (_options.NormalizeLevels && block.Level.HasValue) {
                if (string.Equals(block.Kind, "heading", StringComparison.OrdinalIgnoreCase)) {
                    block.Level = Math.Max(1, Math.Min(6, block.Level.Value));
                } else if ((block.Kind?.IndexOf("list", StringComparison.OrdinalIgnoreCase) ?? -1) >= 0) {
                    block.Level = Math.Max(1, block.Level.Value);
                }
            }
        }
        return document;
    }
}
