using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>Combined source recovery and SVG image-rendering evidence for an export operation.</summary>
public sealed class PublisherConversionReport : IOfficeConversionReport {
    internal PublisherConversionReport(PublisherReadReport source, IEnumerable<OfficeImageExportDiagnostic> images) {
        ReadReport = source;
        FidelityDiagnostics = Array.AsReadOnly(source.FidelityDiagnostics.Concat(images.Select(item =>
            new OfficeConversionFidelityDiagnostic(item.Code, item.Message, item.LossKind, "OfficeIMO.Publisher.Svg", item.Source))).ToArray());
    }
    /// <summary>Immutable source decoding and reconstruction evidence carried into this operation.</summary>
    public PublisherReadReport ReadReport { get; }
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Publisher SVG conversion reported possible fidelity loss. Inspect FidelityDiagnostics.", this);
    }
}
