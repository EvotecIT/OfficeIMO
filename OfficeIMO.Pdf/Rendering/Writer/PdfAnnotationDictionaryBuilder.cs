using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfAnnotationDictionaryBuilder {
    internal static string BuildUriLinkAnnotation(double x1, double y1, double x2, double y2, string uri, string? contents = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.AbsoluteUri(uri, nameof(uri));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /URI /URI " +
            PdfSyntaxEscaper.LiteralString(uri) +
            " >> >>\n";
    }

    internal static string BuildGoToNamedDestinationLinkAnnotation(double x1, double y1, double x2, double y2, string destinationName, string? contents = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(destinationName, nameof(destinationName));

        return "<< /Type /Annot /Subtype /Link /Border [0 0 0]" + BuildContentsEntry(contents) + " /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /A << /S /GoTo /D " +
            PdfSyntaxEscaper.LiteralString(destinationName) +
            " >> >>\n";
    }

    private static string BuildContentsEntry(string? contents) =>
        string.IsNullOrWhiteSpace(contents)
            ? string.Empty
            : " /Contents " + PdfSyntaxEscaper.LiteralString(contents!);

    private static void ValidateRectangle(double x1, double y1, double x2, double y2) {
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateFinite(x2, nameof(x2));
        ValidateFinite(y2, nameof(y2));

        if (x2 <= x1) {
            throw new ArgumentOutOfRangeException(nameof(x2), x2, "PDF link annotation rectangle must have positive width.");
        }

        if (y2 <= y1) {
            throw new ArgumentOutOfRangeException(nameof(y2), y2, "PDF link annotation rectangle must have positive height.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF annotation coordinates must be finite numbers.");
        }
    }

    private static string FormatCoordinate(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);
}
