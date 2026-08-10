using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

/// <summary>Contains one native editable Visio projection and its semantic fidelity report.</summary>
public sealed class OfficeVisioVisualConversionResult {
    internal OfficeVisioVisualConversionResult(
        VisualArtifactInterchangeEnvelope envelope,
        VisioDocument document,
        VisioPage page,
        OfficeVisioVisualConversionReport report) {
        Envelope = envelope;
        Document = document;
        Page = page;
        Report = report;
    }

    /// <summary>Gets the validated CFX semantic envelope used for projection.</summary>
    public VisualArtifactInterchangeEnvelope Envelope { get; }

    /// <summary>Gets the editable Visio document.</summary>
    public VisioDocument Document { get; }

    /// <summary>Gets the generated editable Visio page.</summary>
    public VisioPage Page { get; }

    /// <summary>Gets the semantic fidelity report.</summary>
    public OfficeVisioVisualConversionReport Report { get; }
}
