namespace OfficeIMO.Rtf;

/// <summary>Pairs a converted value with its shared RTF conversion report.</summary>
public sealed class RtfConversionResult<T> {
    /// <summary>Initializes a conversion result.</summary>
    public RtfConversionResult(T value, RtfConversionReport report) {
        Value = value;
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Converted value.</summary>
    public T Value { get; }

    /// <summary>Fidelity and policy report.</summary>
    public RtfConversionReport Report { get; }

    /// <summary>Requires a lossless conversion and returns the converted value.</summary>
    public T RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
