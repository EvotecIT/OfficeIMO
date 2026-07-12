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

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !Report.Diagnostics.Any(static diagnostic => diagnostic.Severity == RtfConversionSeverity.Error);

    /// <summary>Whether any source feature was flattened, omitted, blocked, or failed.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the value or throws when conversion reported an error.</summary>
    public T RequireValue() {
        if (!Succeeded) throw new RtfConversionLossException(Report);
        return Value;
    }

    /// <summary>Requires a lossless conversion and returns the converted value.</summary>
    public T RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
