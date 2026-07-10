namespace OfficeIMO.Rtf;

/// <summary>Thrown by strict conversion workflows when fidelity loss is reported.</summary>
public sealed class RtfConversionLossException : InvalidOperationException {
    /// <summary>Initializes a strict-conversion exception.</summary>
    public RtfConversionLossException(RtfConversionReport report)
        : base("RTF conversion reported fidelity loss or a blocked feature.") {
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Report that caused strict conversion to fail.</summary>
    public RtfConversionReport Report { get; }
}
