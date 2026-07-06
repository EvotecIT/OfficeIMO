namespace OfficeIMO.Pdf;

internal readonly struct PdfAnnotationPathPoint {
    internal PdfAnnotationPathPoint(double x, double y) {
        X = x;
        Y = y;
    }

    internal double X { get; }

    internal double Y { get; }
}

internal static partial class PdfAnnotationDictionaryBuilder {
    internal static string BuildInkAppearanceContent(double width, double height, IReadOnlyList<IReadOnlyList<PdfAnnotationPathPoint>> paths, PdfColor? strokeColor = null, double borderWidth = 1D, IReadOnlyList<double>? borderDashPattern = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(paths, nameof(paths));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        if (borderWidth <= 0D || !HasRenderablePath(paths)) {
            return "q\nQ\n";
        }

        PdfColor resolvedStrokeColor = strokeColor ?? PdfColor.Black;
        var builder = new StringBuilder();
        builder.Append("q\n")
            .Append(BuildBorderStrokeOperators(resolvedStrokeColor, borderWidth, borderDashPattern))
            .Append("1 J 1 j\n");

        for (int i = 0; i < paths.Count; i++) {
            AppendOpenPath(builder, paths[i]);
        }

        builder.Append("Q\n");
        return builder.ToString();
    }

    internal static string BuildPathAnnotationAppearanceContent(double width, double height, string subtype, IReadOnlyList<PdfAnnotationPathPoint> vertices, PdfColor? strokeColor = null, PdfColor? fillColor = null, double borderWidth = 1D, IReadOnlyList<double>? borderDashPattern = null, string? startEnding = null, string? endEnding = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(vertices, nameof(vertices));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        ValidatePathAnnotationSubtype(subtype);

        bool isPolygon = string.Equals(subtype, "Polygon", StringComparison.Ordinal);
        bool hasStroke = borderWidth > 0D;
        if (vertices.Count < 2 || (!hasStroke && (!isPolygon || !fillColor.HasValue))) {
            return "q\nQ\n";
        }

        PdfColor resolvedStrokeColor = strokeColor ?? PdfColor.Black;
        PdfColor resolvedFillColor = fillColor ?? resolvedStrokeColor;
        var builder = new StringBuilder();
        builder.Append("q\n");
        if (fillColor.HasValue && isPolygon) {
            builder.Append(FormatColor(fillColor.Value)).Append(" rg ");
        }

        if (hasStroke) {
            builder.Append(BuildBorderStrokeOperators(resolvedStrokeColor, borderWidth, borderDashPattern));
        }

        AppendVerticesPath(builder, vertices, isPolygon);
        if (isPolygon) {
            builder.Append(fillColor.HasValue && hasStroke ? "B\n" : fillColor.HasValue ? "f\n" : "S\n");
        } else if (hasStroke) {
            builder.Append("S\n");
            AppendPolylineEndings(builder, vertices, borderWidth, resolvedStrokeColor, resolvedFillColor, startEnding, endEnding);
        }

        builder.Append("Q\n");
        return builder.ToString();
    }

    private static bool HasRenderablePath(IReadOnlyList<IReadOnlyList<PdfAnnotationPathPoint>> paths) {
        for (int i = 0; i < paths.Count; i++) {
            if (paths[i].Count >= 2) {
                return true;
            }
        }

        return false;
    }

    private static void AppendOpenPath(StringBuilder builder, IReadOnlyList<PdfAnnotationPathPoint> points) {
        if (points.Count < 2) {
            return;
        }

        builder.Append(FormatCoordinate(points[0].X)).Append(' ').Append(FormatCoordinate(points[0].Y)).Append(" m ");
        for (int i = 1; i < points.Count; i++) {
            builder.Append(FormatCoordinate(points[i].X)).Append(' ').Append(FormatCoordinate(points[i].Y)).Append(" l ");
        }

        builder.Append("S\n");
    }

    private static void AppendVerticesPath(StringBuilder builder, IReadOnlyList<PdfAnnotationPathPoint> vertices, bool closePath) {
        builder.Append(FormatCoordinate(vertices[0].X)).Append(' ').Append(FormatCoordinate(vertices[0].Y)).Append(" m ");
        for (int i = 1; i < vertices.Count; i++) {
            builder.Append(FormatCoordinate(vertices[i].X)).Append(' ').Append(FormatCoordinate(vertices[i].Y)).Append(" l ");
        }

        if (closePath) {
            builder.Append("h ");
        }
    }

    private static void AppendPolylineEndings(StringBuilder builder, IReadOnlyList<PdfAnnotationPathPoint> vertices, double borderWidth, PdfColor strokeColor, PdfColor fillColor, string? startEnding, string? endEnding) {
        if (vertices.Count < 2) {
            return;
        }

        AppendLineEnding(builder, startEnding, vertices[0].X, vertices[0].Y, vertices[1].X, vertices[1].Y, borderWidth, strokeColor, fillColor);
        int last = vertices.Count - 1;
        AppendLineEnding(builder, endEnding, vertices[last].X, vertices[last].Y, vertices[last - 1].X, vertices[last - 1].Y, borderWidth, strokeColor, fillColor);
    }

    private static void ValidatePathAnnotationSubtype(string subtype) {
        if (!string.Equals(subtype, "Polygon", StringComparison.Ordinal) &&
            !string.Equals(subtype, "PolyLine", StringComparison.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(subtype), subtype, "PDF path annotation subtype must be Polygon or PolyLine.");
        }
    }
}
