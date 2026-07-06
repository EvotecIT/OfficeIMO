namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private void ReadAnnotationVisualStyleMetadata(
        PdfDictionary annotation,
        string subtype,
        double width,
        double height,
        out IReadOnlyList<double> interiorColor,
        out double? opacity,
        out double? borderWidth,
        out string? borderStyle,
        out IReadOnlyList<double> borderDashPattern,
        out string? borderEffectStyle,
        out double? borderEffectIntensity,
        out IReadOnlyList<double> rectangleDifferences,
        out IReadOnlyList<double> calloutLine,
        out string? calloutLineEnding,
        out string? lineStartEnding,
        out string? lineEndEnding) {
        interiorColor = ReadNumberArray(annotation.Items.TryGetValue("IC", out PdfObject? interiorColorObject) ? interiorColorObject : null);
        opacity = TryReadAnnotationOpacity(annotation);
        borderWidth = TryReadAnnotationBorderWidth(annotation);
        borderStyle = TryReadAnnotationBorderStyle(annotation);
        borderDashPattern = TryReadAnnotationBorderDashPattern(annotation) ?? Array.Empty<double>();
        TryReadAnnotationBorderEffect(annotation, out borderEffectStyle, out borderEffectIntensity);
        rectangleDifferences = string.Equals(subtype, "FreeText", StringComparison.Ordinal)
            ? TryReadFreeTextRectangleDifferences(annotation, width, height) ?? Array.Empty<double>()
            : Array.Empty<double>();
        calloutLine = string.Equals(subtype, "FreeText", StringComparison.Ordinal)
            ? TryReadFreeTextCalloutLine(annotation) ?? Array.Empty<double>()
            : Array.Empty<double>();
        calloutLineEnding = string.Equals(subtype, "FreeText", StringComparison.Ordinal)
            ? TryReadFreeTextCalloutLineEnding(annotation)
            : null;
        TryReadLineEndings(annotation, out lineStartEnding, out lineEndEnding);
    }

    private double? TryReadAnnotationOpacity(PdfDictionary annotation) {
        if (!annotation.Items.TryGetValue("CA", out PdfObject? opacityObject) ||
            ResolveObject(opacityObject) is not PdfNumber opacity ||
            !IsFiniteRange(opacity.Value, 0D, 1D)) {
            return null;
        }

        return opacity.Value;
    }

    private double? TryReadAnnotationBorderWidth(PdfDictionary annotation) {
        PdfDictionary? borderStyle = ResolveDictionary(annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null);
        if (borderStyle != null &&
            borderStyle.Items.TryGetValue("W", out PdfObject? borderStyleWidthObject) &&
            TryReadNonNegativeFiniteNumber(borderStyleWidthObject, out double borderStyleWidth)) {
            return borderStyleWidth;
        }

        PdfArray? border = ResolveArray(annotation.Items.TryGetValue("Border", out PdfObject? borderObject) ? borderObject : null);
        if (border != null &&
            border.Items.Count >= 3 &&
            TryReadNonNegativeFiniteNumber(border.Items[2], out double borderWidth)) {
            return borderWidth;
        }

        return null;
    }

    private string? TryReadAnnotationBorderStyle(PdfDictionary annotation) {
        PdfDictionary? borderStyle = ResolveDictionary(annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null);
        if (borderStyle?.Get<PdfName>("S") is not PdfName styleName) {
            return null;
        }

        switch (styleName.Name) {
            case "S":
                return "Solid";
            case "D":
                return "Dashed";
            case "U":
                return "Underline";
            case "B":
                return "Beveled";
            case "I":
                return "Inset";
            default:
                return styleName.Name;
        }
    }

    private double[]? TryReadAnnotationBorderDashPattern(PdfDictionary annotation) {
        PdfDictionary? borderStyle = ResolveDictionary(annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null);
        if (borderStyle?.Get<PdfName>("S")?.Name != "D") {
            return null;
        }

        if (!borderStyle.Items.TryGetValue("D", out PdfObject? dashObject)) {
            return new[] { 3D };
        }

        return TryReadDashPattern(dashObject);
    }

    private double[]? TryReadDashPattern(PdfObject dashObject) {
        PdfArray? dashArray = ResolveArray(dashObject);
        if (dashArray == null || dashArray.Items.Count == 0) {
            return null;
        }

        var values = new double[dashArray.Items.Count];
        bool hasPositiveSegment = false;
        for (int i = 0; i < dashArray.Items.Count; i++) {
            if (!TryReadNonNegativeFiniteNumber(dashArray.Items[i], out double segment)) {
                return null;
            }

            if (segment > 0D) {
                hasPositiveSegment = true;
            }

            values[i] = segment;
        }

        return hasPositiveSegment ? values : null;
    }

    private void TryReadAnnotationBorderEffect(PdfDictionary annotation, out string? style, out double? intensity) {
        style = null;
        intensity = null;
        PdfDictionary? borderEffect = ResolveDictionary(annotation.Items.TryGetValue("BE", out PdfObject? borderEffectObject) ? borderEffectObject : null);
        if (borderEffect?.Get<PdfName>("S") is not PdfName styleName) {
            return;
        }

        style = string.Equals(styleName.Name, "C", StringComparison.Ordinal)
            ? "Cloudy"
            : styleName.Name;
        if (borderEffect.Items.TryGetValue("I", out PdfObject? intensityObject) &&
            ResolveObject(intensityObject) is PdfNumber rawIntensity &&
            rawIntensity.Value > 0D &&
            !double.IsNaN(rawIntensity.Value) &&
            !double.IsInfinity(rawIntensity.Value)) {
            intensity = Math.Min(2D, rawIntensity.Value);
        } else if (string.Equals(styleName.Name, "C", StringComparison.Ordinal)) {
            intensity = 1D;
        }
    }

    private double[]? TryReadFreeTextRectangleDifferences(PdfDictionary annotation, double width, double height) {
        PdfArray? rectangleDifferences = ResolveArray(annotation.Items.TryGetValue("RD", out PdfObject? rectangleDifferencesObject) ? rectangleDifferencesObject : null);
        if (rectangleDifferences == null ||
            rectangleDifferences.Items.Count < 4 ||
            !TryReadNonNegativeFiniteNumber(rectangleDifferences.Items[0], out double left) ||
            !TryReadNonNegativeFiniteNumber(rectangleDifferences.Items[1], out double top) ||
            !TryReadNonNegativeFiniteNumber(rectangleDifferences.Items[2], out double right) ||
            !TryReadNonNegativeFiniteNumber(rectangleDifferences.Items[3], out double bottom) ||
            left + right >= width ||
            top + bottom >= height) {
            return null;
        }

        return new[] { left, top, right, bottom };
    }

    private double[]? TryReadFreeTextCalloutLine(PdfDictionary annotation) {
        PdfArray? calloutLine = ResolveArray(annotation.Items.TryGetValue("CL", out PdfObject? calloutObject) ? calloutObject : null);
        if (calloutLine == null ||
            (calloutLine.Items.Count != 4 && calloutLine.Items.Count != 6)) {
            return null;
        }

        var values = new double[calloutLine.Items.Count];
        for (int i = 0; i < calloutLine.Items.Count; i++) {
            if (ResolveObject(calloutLine.Items[i]) is not PdfNumber coordinate ||
                !IsFinite(coordinate.Value)) {
                return null;
            }

            values[i] = coordinate.Value;
        }

        return values;
    }

    private string? TryReadFreeTextCalloutLineEnding(PdfDictionary annotation) {
        if (!annotation.Items.TryGetValue("LE", out PdfObject? lineEndingObject)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(lineEndingObject);
        if (resolved is PdfName lineEndingName) {
            return lineEndingName.Name;
        }

        if (resolved is PdfArray lineEndings &&
            lineEndings.Items.Count > 0 &&
            ResolveObject(lineEndings.Items[0]) is PdfName firstEndingName) {
            return firstEndingName.Name;
        }

        return null;
    }

    private void TryReadLineEndings(PdfDictionary annotation, out string? startEnding, out string? endEnding) {
        startEnding = null;
        endEnding = null;
        if (!annotation.Items.TryGetValue("LE", out PdfObject? lineEndingObject) ||
            ResolveObject(lineEndingObject) is not PdfArray lineEndings ||
            lineEndings.Items.Count < 2) {
            return;
        }

        if (ResolveObject(lineEndings.Items[0]) is PdfName startName) {
            startEnding = startName.Name;
        }

        if (ResolveObject(lineEndings.Items[1]) is PdfName endName) {
            endEnding = endName.Name;
        }
    }

    private bool TryReadNonNegativeFiniteNumber(PdfObject numberObject, out double value) {
        value = 0D;
        if (ResolveObject(numberObject) is not PdfNumber number ||
            number.Value < 0D ||
            !IsFinite(number.Value)) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private static bool IsFiniteRange(double value, double minimum, double maximum) =>
        IsFinite(value) && value >= minimum && value <= maximum;
}
