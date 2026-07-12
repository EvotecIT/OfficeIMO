namespace OfficeIMO.Html;

/// <summary>
/// Shared result contract for HTML conversions that produce a native target artifact.
/// </summary>
/// <typeparam name="T">Value produced by the conversion.</typeparam>
public abstract class HtmlConversionResult<T> {
    /// <summary>Creates a conversion result for the supplied value.</summary>
    protected HtmlConversionResult(T value) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Value produced by the conversion.</summary>
    public T Value { get; }

    /// <summary>Structured conversion diagnostics in emission order.</summary>
    public HtmlDiagnosticReport Diagnostics { get; } = new HtmlDiagnosticReport();

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !Diagnostics.HasErrors;

    /// <summary>Whether the conversion approximated, omitted, or failed any source content.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.LossKind != HtmlConversionLossKind.None);

    /// <summary>
    /// Returns the native artifact when conversion succeeded, or throws a structured conversion exception.
    /// </summary>
    public T RequireValue() {
        if (!Succeeded) {
            throw new HtmlConversionException(Diagnostics.Diagnostics);
        }

        return Value;
    }
}
