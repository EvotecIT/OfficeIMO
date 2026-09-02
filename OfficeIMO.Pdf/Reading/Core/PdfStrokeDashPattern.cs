namespace OfficeIMO.Pdf;

/// <summary>Retains the exact PDF dash array and phase that affect stroke rendering.</summary>
internal readonly struct PdfStrokeDashPattern {
    internal PdfStrokeDashPattern(IReadOnlyList<double> array, double phase) {
        Array = array.Count == 0 ? System.Array.Empty<double>() : array.ToArray();
        Phase = phase;
    }

    internal IReadOnlyList<double> Array { get; }

    internal double Phase { get; }

    internal static PdfStrokeDashPattern Solid => new PdfStrokeDashPattern(System.Array.Empty<double>(), 0D);
}
