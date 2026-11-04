namespace OfficeIMO.Rtf;

/// <summary>Thrown by strict conversion workflows when fidelity loss is reported.</summary>
public sealed class RtfConversionLossException : OfficeConversionException {
    /// <summary>Initializes a strict-conversion exception.</summary>
    public RtfConversionLossException(RtfConversionReport report)
        : base("RTF conversion reported fidelity loss or a blocked feature.", report) {
        Report = report;
    }

    /// <summary>Report that caused strict conversion to fail.</summary>
    public new RtfConversionReport Report { get; }
}
