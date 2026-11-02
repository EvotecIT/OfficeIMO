namespace OfficeIMO.OpenDocument;

/// <summary>Converted document together with the feature mapping evidence for that conversion.</summary>
public sealed class OdfConversionResult<TDocument> : OfficeConversionResult<TDocument> where TDocument : class {
    /// <summary>Gets the typed conversion report.</summary>
    public new OdfConversionReport Report => (OdfConversionReport)base.Report;

    /// <summary>Creates a conversion result.</summary>
    public OdfConversionResult(TDocument value, OdfConversionReport report) : base(value, report) { }


    /// <summary>Returns the converted document after applying an explicit loss policy.</summary>
    public TDocument GetValue(OdfConversionLossPolicy policy) {
        try {
            Report.ApplyPolicy(policy);
            return Value;
        } catch {
            if (Value is IDisposable disposable) disposable.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Applies a loss policy and returns this result for fluent adapter composition. A rejected disposable
    /// conversion value is disposed before the policy exception is rethrown.
    /// </summary>
    public OdfConversionResult<TDocument> ApplyPolicy(OdfConversionLossPolicy policy) {
        try {
            Report.ApplyPolicy(policy);
            return this;
        } catch {
            if (Value is IDisposable disposable) disposable.Dispose();
            throw;
        }
    }
}
