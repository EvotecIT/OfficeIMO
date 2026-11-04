namespace OfficeIMO.Rtf;

/// <summary>Pairs a converted value with its shared RTF conversion report.</summary>
public sealed class RtfConversionResult<T> : OfficeConversionResult<T> where T : class {
    /// <summary>Gets the typed conversion report.</summary>
    public new RtfConversionReport Report => (RtfConversionReport)base.Report;

    /// <summary>Initializes a conversion result.</summary>
    public RtfConversionResult(T value, RtfConversionReport report) : base(value, report) { }

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !Report.NativeDiagnostics.Any(static diagnostic => diagnostic.Severity == RtfConversionSeverity.Error);

    /// <summary>Returns the value or throws when conversion reported an error.</summary>
    public T RequireValue() {
        if (!Succeeded) throw new RtfConversionLossException(Report);
        return Value;
    }

}
