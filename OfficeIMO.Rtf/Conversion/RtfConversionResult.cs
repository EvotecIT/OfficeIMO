namespace OfficeIMO.Rtf;

/// <summary>Pairs a converted value with its shared RTF conversion report.</summary>
public sealed class RtfConversionResult<T> : OfficeConversionResult<T, RtfConversionReport> where T : class {
    /// <summary>Initializes a conversion result.</summary>
    public RtfConversionResult(T value, RtfConversionReport report) : base(value, report) { }

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public override bool Succeeded => !Report.Diagnostics.Any(static diagnostic => diagnostic.Severity == RtfConversionSeverity.Error);

    /// <summary>Returns the value or throws when conversion reported an error.</summary>
    public override T RequireValue() {
        if (!Succeeded) throw new RtfConversionLossException(Report);
        return base.RequireValue();
    }

    /// <summary>Requires a lossless conversion and returns the converted value.</summary>
    public override T RequireNoLoss() {
        Report.RequireNoLoss();
        return base.RequireValue();
    }
}
