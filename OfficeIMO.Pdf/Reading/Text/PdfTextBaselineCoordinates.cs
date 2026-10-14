namespace OfficeIMO.Pdf;

/// <summary>Projects page-space span origins along and across a painted text baseline without changing source geometry.</summary>
internal readonly struct PdfTextBaselineCoordinates {
    private readonly double _cos;
    private readonly double _sin;

    internal PdfTextBaselineCoordinates(double rotationDegrees) {
        double radians = rotationDegrees * Math.PI / 180D;
        _cos = Math.Cos(radians);
        _sin = Math.Sin(radians);
    }

    internal double Along(PdfTextSpan span) => _cos * span.X + _sin * span.Y;
    internal double Normal(PdfTextSpan span) => -_sin * span.X + _cos * span.Y;
}
