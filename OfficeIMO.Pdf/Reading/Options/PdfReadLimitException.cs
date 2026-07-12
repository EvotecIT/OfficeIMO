namespace OfficeIMO.Pdf;

/// <summary>Thrown when a PDF exceeds an explicit bounded-read resource limit.</summary>
public sealed class PdfReadLimitException : IOException {
    internal PdfReadLimitException(PdfReadLimitKind kind, long limit, long actual, string message)
        : base(message) {
        Kind = kind;
        Limit = limit;
        Actual = actual;
    }

    /// <summary>Resource budget that was exceeded.</summary>
    public PdfReadLimitKind Kind { get; }

    /// <summary>Configured maximum value.</summary>
    public long Limit { get; }

    /// <summary>Observed value when the parser stopped.</summary>
    public long Actual { get; }

    internal static PdfReadLimitException Create(PdfReadLimitKind kind, long limit, long actual) =>
        new(
            kind,
            limit,
            actual,
            "PDF read limit exceeded for " + kind + ": observed " + actual.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            ", maximum " + limit.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
}
