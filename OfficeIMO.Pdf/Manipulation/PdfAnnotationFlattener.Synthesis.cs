namespace OfficeIMO.Pdf;

public static partial class PdfAnnotationFlattener {
    private static readonly char[] AppearanceTokenSeparators = { ' ', '\t', '\r', '\n' };

    private static PdfReference CreateSyntheticAppearanceReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        string subtype,
        double x,
        double y,
        double width,
        double height,
        ref int nextObjectNumber) {
        string content;
        bool needsHelvetica;

        if (string.Equals(subtype, "FreeText", StringComparison.Ordinal)) {
            content = BuildSyntheticFreeTextAppearance(objects, annotation, width, height);
            needsHelvetica = true;
        } else if (string.Equals(subtype, "Highlight", StringComparison.Ordinal)) {
            content = BuildSyntheticHighlightAppearance(objects, annotation, x, y, width, height);
            needsHelvetica = false;
        } else if (IsTextMarkupSubtype(subtype)) {
            content = BuildSyntheticTextMarkupAppearance(objects, annotation, subtype, x, y, width, height);
            needsHelvetica = false;
        } else if (IsShapeSubtype(subtype)) {
            content = BuildSyntheticShapeAppearance(objects, annotation, subtype, width, height);
            needsHelvetica = false;
        } else if (string.Equals(subtype, "Line", StringComparison.Ordinal)) {
            content = BuildSyntheticLineAppearance(objects, annotation, x, y, width, height);
            needsHelvetica = false;
        } else if (string.Equals(subtype, "Ink", StringComparison.Ordinal)) {
            content = BuildSyntheticInkAppearance(objects, annotation, x, y, width, height);
            needsHelvetica = false;
        } else if (IsPathAnnotationSubtype(subtype)) {
            content = BuildSyntheticPathAnnotationAppearance(objects, annotation, subtype, x, y, width, height);
            needsHelvetica = false;
        } else if (string.Equals(subtype, "Stamp", StringComparison.Ordinal)) {
            content = BuildSyntheticStampAppearance(objects, annotation, width, height);
            needsHelvetica = true;
        } else if (string.Equals(subtype, "Caret", StringComparison.Ordinal)) {
            content = BuildSyntheticCaretAppearance(objects, annotation, width, height);
            needsHelvetica = false;
        } else {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }

        int appearanceObjectNumber = nextObjectNumber++;
        objects[appearanceObjectNumber] = new PdfIndirectObject(
            appearanceObjectNumber,
            0,
            CreateSyntheticAppearanceStream(width, height, content, needsHelvetica, string.Equals(subtype, "Highlight", StringComparison.Ordinal)));
        return new PdfReference(appearanceObjectNumber, 0);
    }

    private static string BuildSyntheticFreeTextAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double width, double height) {
        string contents = TryReadString(objects, annotation, "Contents") ?? string.Empty;
        double fontSize = TryReadFreeTextFontSize(objects, annotation) ?? 10D;
        PdfColor? textColor = TryReadFreeTextTextColor(objects, annotation);
        PdfColor? borderColor = TryReadColor(objects, annotation, "C");
        double borderWidth = TryReadBorderWidth(objects, annotation);
        PdfColor? fillColor = TryReadColor(objects, annotation, "IC");
        PdfAlign textAlign = TryReadFreeTextAlignment(objects, annotation);

        return PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(
            width,
            height,
            contents,
            fontSize,
            textColor,
            borderColor,
            borderWidth,
            fillColor,
            textAlign);
    }

    private static string BuildSyntheticHighlightAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        IReadOnlyList<PdfHighlightQuad> quads = TryReadHighlightQuadPoints(objects, annotation, x, y, width, height);
        return quads.Count > 0
            ? PdfAnnotationDictionaryBuilder.BuildHighlightAppearanceContent(width, height, quads, TryReadColor(objects, annotation, "C"))
            : PdfAnnotationDictionaryBuilder.BuildHighlightAppearanceContent(width, height, TryReadColor(objects, annotation, "C"));
    }

    private static string BuildSyntheticTextMarkupAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, string subtype, double x, double y, double width, double height) {
        IReadOnlyList<PdfHighlightQuad> quads = TryReadHighlightQuadPoints(objects, annotation, x, y, width, height);
        return quads.Count > 0
            ? PdfAnnotationDictionaryBuilder.BuildTextMarkupAppearanceContent(width, height, quads, subtype, TryReadColor(objects, annotation, "C"))
            : PdfAnnotationDictionaryBuilder.BuildTextMarkupAppearanceContent(width, height, subtype, TryReadColor(objects, annotation, "C"));
    }

    private static bool IsTextMarkupSubtype(string subtype) {
        return string.Equals(subtype, "Underline", StringComparison.Ordinal) ||
            string.Equals(subtype, "StrikeOut", StringComparison.Ordinal) ||
            string.Equals(subtype, "Squiggly", StringComparison.Ordinal);
    }

    private static string BuildSyntheticShapeAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, string subtype, double width, double height) {
        return PdfAnnotationDictionaryBuilder.BuildShapeAppearanceContent(
            width,
            height,
            subtype,
            TryReadColor(objects, annotation, "C"),
            TryReadColor(objects, annotation, "IC"),
            TryReadBorderWidth(objects, annotation));
    }

    private static bool IsShapeSubtype(string subtype) {
        return string.Equals(subtype, "Square", StringComparison.Ordinal) ||
            string.Equals(subtype, "Circle", StringComparison.Ordinal);
    }

    private static string BuildSyntheticInkAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        IReadOnlyList<IReadOnlyList<PdfAnnotationPathPoint>> paths = TryReadInkList(objects, annotation, x, y, width, height);
        if (paths.Count == 0) {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }

        return PdfAnnotationDictionaryBuilder.BuildInkAppearanceContent(
            width,
            height,
            paths,
            TryReadColor(objects, annotation, "C"),
            TryReadBorderWidth(objects, annotation));
    }

    private static string BuildSyntheticPathAnnotationAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, string subtype, double x, double y, double width, double height) {
        IReadOnlyList<PdfAnnotationPathPoint> vertices = TryReadVertices(objects, annotation, x, y, width, height);
        if (vertices.Count < 2) {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }

        TryReadLineEndings(objects, annotation, out string? startEnding, out string? endEnding);
        return PdfAnnotationDictionaryBuilder.BuildPathAnnotationAppearanceContent(
            width,
            height,
            subtype,
            vertices,
            TryReadColor(objects, annotation, "C"),
            TryReadColor(objects, annotation, "IC"),
            TryReadBorderWidth(objects, annotation),
            startEnding,
            endEnding);
    }

    private static bool IsPathAnnotationSubtype(string subtype) {
        return string.Equals(subtype, "Polygon", StringComparison.Ordinal) ||
            string.Equals(subtype, "PolyLine", StringComparison.Ordinal);
    }

    private static string BuildSyntheticStampAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double width, double height) {
        string stampName = TryReadName(objects, annotation, "Name") ?? "Stamp";
        return PdfAnnotationDictionaryBuilder.BuildStampAppearanceContent(
            width,
            height,
            stampName,
            TryReadColor(objects, annotation, "C"),
            TryReadColor(objects, annotation, "IC"),
            TryReadBorderWidth(objects, annotation));
    }

    private static string BuildSyntheticCaretAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double width, double height) {
        return PdfAnnotationDictionaryBuilder.BuildCaretAppearanceContent(
            width,
            height,
            TryReadColor(objects, annotation, "C"),
            TryReadBorderWidth(objects, annotation));
    }

    private static string BuildSyntheticLineAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        if (!TryReadLineCoordinates(objects, annotation, x, y, width, height, out double x1, out double y1, out double x2, out double y2)) {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }

        TryReadLineEndings(objects, annotation, out string? startEnding, out string? endEnding);
        return PdfAnnotationDictionaryBuilder.BuildLineAppearanceContent(
            width,
            height,
            x1,
            y1,
            x2,
            y2,
            TryReadColor(objects, annotation, "C"),
            TryReadColor(objects, annotation, "IC"),
            TryReadBorderWidth(objects, annotation),
            startEnding,
            endEnding);
    }

    private static bool TryReadLineCoordinates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height, out double x1, out double y1, out double x2, out double y2) {
        x1 = 0D;
        y1 = 0D;
        x2 = 0D;
        y2 = 0D;
        if (!annotation.Items.TryGetValue("L", out var lineObject) ||
            ResolveObject(objects, lineObject) is not PdfArray line ||
            line.Items.Count < 4 ||
            ResolveObject(objects, line.Items[0]) is not PdfNumber rawX1 ||
            ResolveObject(objects, line.Items[1]) is not PdfNumber rawY1 ||
            ResolveObject(objects, line.Items[2]) is not PdfNumber rawX2 ||
            ResolveObject(objects, line.Items[3]) is not PdfNumber rawY2) {
            return false;
        }

        x1 = ClampCoordinate(rawX1.Value - x, 0D, width);
        y1 = ClampCoordinate(rawY1.Value - y, 0D, height);
        x2 = ClampCoordinate(rawX2.Value - x, 0D, width);
        y2 = ClampCoordinate(rawY2.Value - y, 0D, height);
        return Math.Abs(x1 - x2) > 0.0001D || Math.Abs(y1 - y2) > 0.0001D;
    }

    private static void TryReadLineEndings(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, out string? startEnding, out string? endEnding) {
        startEnding = null;
        endEnding = null;
        if (!annotation.Items.TryGetValue("LE", out var lineEndingObject) ||
            ResolveObject(objects, lineEndingObject) is not PdfArray lineEndings ||
            lineEndings.Items.Count < 2) {
            return;
        }

        if (ResolveObject(objects, lineEndings.Items[0]) is PdfName startName) {
            startEnding = startName.Name;
        }

        if (ResolveObject(objects, lineEndings.Items[1]) is PdfName endName) {
            endEnding = endName.Name;
        }
    }

    private static IReadOnlyList<IReadOnlyList<PdfAnnotationPathPoint>> TryReadInkList(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        if (!annotation.Items.TryGetValue("InkList", out var inkListObject) ||
            ResolveObject(objects, inkListObject) is not PdfArray inkList ||
            inkList.Items.Count == 0) {
            return Array.Empty<IReadOnlyList<PdfAnnotationPathPoint>>();
        }

        var paths = new List<IReadOnlyList<PdfAnnotationPathPoint>>();
        for (int i = 0; i < inkList.Items.Count; i++) {
            if (ResolveObject(objects, inkList.Items[i]) is not PdfArray pathArray) {
                return Array.Empty<IReadOnlyList<PdfAnnotationPathPoint>>();
            }

            List<PdfAnnotationPathPoint> points = ReadPathPoints(objects, pathArray, x, y, width, height);
            if (points.Count >= 2) {
                paths.Add(points);
            }
        }

        return paths;
    }

    private static IReadOnlyList<PdfAnnotationPathPoint> TryReadVertices(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        if (!annotation.Items.TryGetValue("Vertices", out var verticesObject) ||
            ResolveObject(objects, verticesObject) is not PdfArray vertices) {
            return Array.Empty<PdfAnnotationPathPoint>();
        }

        return ReadPathPoints(objects, vertices, x, y, width, height);
    }

    private static List<PdfAnnotationPathPoint> ReadPathPoints(Dictionary<int, PdfIndirectObject> objects, PdfArray array, double x, double y, double width, double height) {
        if (array.Items.Count < 4 || array.Items.Count % 2 != 0) {
            return new List<PdfAnnotationPathPoint>();
        }

        var points = new List<PdfAnnotationPathPoint>(array.Items.Count / 2);
        for (int i = 0; i < array.Items.Count; i += 2) {
            if (ResolveObject(objects, array.Items[i]) is not PdfNumber rawX ||
                ResolveObject(objects, array.Items[i + 1]) is not PdfNumber rawY) {
                return new List<PdfAnnotationPathPoint>();
            }

            points.Add(new PdfAnnotationPathPoint(
                ClampCoordinate(rawX.Value - x, 0D, width),
                ClampCoordinate(rawY.Value - y, 0D, height)));
        }

        return points;
    }

    private static IReadOnlyList<PdfHighlightQuad> TryReadHighlightQuadPoints(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        if (!annotation.Items.TryGetValue("QuadPoints", out var quadPointsObject) ||
            ResolveObject(objects, quadPointsObject) is not PdfArray quadPoints ||
            quadPoints.Items.Count < 8 ||
            quadPoints.Items.Count % 8 != 0) {
            return Array.Empty<PdfHighlightQuad>();
        }

        var quads = new List<PdfHighlightQuad>(quadPoints.Items.Count / 8);
        for (int i = 0; i < quadPoints.Items.Count; i += 8) {
            if (!TryReadQuadPointNumber(objects, quadPoints, i, out double x1) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 1, out double y1) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 2, out double x2) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 3, out double y2) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 4, out double x3) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 5, out double y3) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 6, out double x4) ||
                !TryReadQuadPointNumber(objects, quadPoints, i + 7, out double y4)) {
                return Array.Empty<PdfHighlightQuad>();
            }

            quads.Add(new PdfHighlightQuad(
                ClampCoordinate(x1 - x, 0D, width),
                ClampCoordinate(y1 - y, 0D, height),
                ClampCoordinate(x2 - x, 0D, width),
                ClampCoordinate(y2 - y, 0D, height),
                ClampCoordinate(x3 - x, 0D, width),
                ClampCoordinate(y3 - y, 0D, height),
                ClampCoordinate(x4 - x, 0D, width),
                ClampCoordinate(y4 - y, 0D, height)));
        }

        return quads;
    }

    private static bool TryReadQuadPointNumber(Dictionary<int, PdfIndirectObject> objects, PdfArray quadPoints, int index, out double value) {
        value = 0D;
        if (ResolveObject(objects, quadPoints.Items[index]) is not PdfNumber number) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private static double ClampCoordinate(double value, double min, double max) {
        if (value < min) {
            return min;
        }

        return value > max ? max : value;
    }

    private static PdfStream CreateSyntheticAppearanceStream(double width, double height, string content, bool needsHelvetica, bool usesHighlightBlendMode) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        var bbox = new PdfArray();
        bbox.Items.Add(new PdfNumber(0D));
        bbox.Items.Add(new PdfNumber(0D));
        bbox.Items.Add(new PdfNumber(width));
        bbox.Items.Add(new PdfNumber(height));
        dictionary.Items["BBox"] = bbox;

        if (needsHelvetica || usesHighlightBlendMode) {
            dictionary.Items["Resources"] = CreateSyntheticAppearanceResources(needsHelvetica, usesHighlightBlendMode);
        }

        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfDictionary CreateSyntheticAppearanceResources(bool needsHelvetica, bool usesHighlightBlendMode) {
        var resources = new PdfDictionary();
        if (needsHelvetica) {
            var fonts = new PdfDictionary();
            var helvetica = new PdfDictionary();
            helvetica.Items["Type"] = new PdfName("Font");
            helvetica.Items["Subtype"] = new PdfName("Type1");
            helvetica.Items["BaseFont"] = new PdfName("Helvetica");
            helvetica.Items["Encoding"] = new PdfName("WinAnsiEncoding");
            fonts.Items["Helv"] = helvetica;
            resources.Items["Font"] = fonts;
        }

        if (usesHighlightBlendMode) {
            var extGStates = new PdfDictionary();
            var highlightState = new PdfDictionary();
            highlightState.Items["Type"] = new PdfName("ExtGState");
            highlightState.Items["BM"] = new PdfName("Multiply");
            highlightState.Items["CA"] = new PdfNumber(0.35D);
            highlightState.Items["ca"] = new PdfNumber(0.35D);
            extGStates.Items["OfficeIMOHighlightGs"] = highlightState;
            resources.Items["ExtGState"] = extGStates;
        }

        return resources;
    }

    private static double TryReadBorderWidth(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (!annotation.Items.TryGetValue("Border", out var borderObject) ||
            ResolveObject(objects, borderObject) is not PdfArray border ||
            border.Items.Count < 3 ||
            ResolveObject(objects, border.Items[2]) is not PdfNumber width ||
            width.Value < 0D) {
            return 1D;
        }

        return width.Value;
    }

    private static PdfAlign TryReadFreeTextAlignment(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        double? alignment = TryReadNumber(objects, annotation, "Q");
        if (!alignment.HasValue) {
            return PdfAlign.Left;
        }

        int value = (int)Math.Round(alignment.Value);
        return value == 1
            ? PdfAlign.Center
            : value == 2
                ? PdfAlign.Right
                : PdfAlign.Left;
    }

    private static double? TryReadFreeTextFontSize(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        string? defaultAppearance = TryReadString(objects, annotation, "DA");
        if (defaultAppearance == null || defaultAppearance.Trim().Length == 0) {
            return null;
        }

        string[] tokens = defaultAppearance.Split(AppearanceTokenSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 1; i < tokens.Length; i++) {
            if (string.Equals(tokens[i], "Tf", StringComparison.Ordinal) &&
                double.TryParse(tokens[i - 1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double fontSize) &&
                fontSize > 0D) {
                return fontSize;
            }
        }

        return null;
    }

    private static PdfColor? TryReadFreeTextTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        string? defaultAppearance = TryReadString(objects, annotation, "DA");
        if (defaultAppearance == null || defaultAppearance.Trim().Length == 0) {
            return null;
        }

        string[] tokens = defaultAppearance.Split(AppearanceTokenSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 3; i < tokens.Length; i++) {
            if (!string.Equals(tokens[i], "rg", StringComparison.Ordinal) ||
                !double.TryParse(tokens[i - 3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double r) ||
                !double.TryParse(tokens[i - 2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double g) ||
                !double.TryParse(tokens[i - 1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double b)) {
                continue;
            }

            return new PdfColor(ClampColor(r), ClampColor(g), ClampColor(b));
        }

        return null;
    }
}
