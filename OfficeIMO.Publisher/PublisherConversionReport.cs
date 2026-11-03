using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>Combined source recovery and SVG image-rendering evidence for an export operation.</summary>
public sealed class PublisherConversionReport : OfficeConversionReport {
    internal PublisherConversionReport(PublisherReadReport source, IEnumerable<OfficeImageExportDiagnostic> images) {
        ReadReport = source;
        FidelityDiagnostics = Array.AsReadOnly(source.Diagnostics.Concat(images.Select(item =>
            new global::OfficeIMO.OfficeDiagnostic(item.Code, item.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning, item.Message, item.Source, "OfficeIMO.Publisher.Svg", item.LossKind))).ToArray());
    }
    /// <summary>Immutable source decoding and reconstruction evidence carried into this operation.</summary>
    public PublisherReadReport ReadReport { get; }
    /// <inheritdoc />
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public override bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Publisher SVG conversion reported possible fidelity loss. Inspect Diagnostics.", this);
}
