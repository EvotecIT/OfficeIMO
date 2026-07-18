namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    internal static int RegenerateNormalAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        string subtype = TryReadName(objects, annotation, "Subtype") ?? throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        if (!IsSupportedVisualAnnotation(subtype) || !TryReadRectCoordinates(objects, annotation, out double x, out double y, out double width, out double height)) {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        PdfReference appearance = CreateSyntheticAppearanceReference(objects, annotation, subtype, x, y, width, height, ref nextObjectNumber);
        var appearanceDictionary = new PdfDictionary(); appearanceDictionary.Items["N"] = appearance; annotation.Items["AP"] = appearanceDictionary;
        return appearance.ObjectNumber;
    }

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
        IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources;
        double? opacity = TryReadAnnotationOpacity(objects, annotation);

        if (string.Equals(subtype, "FreeText", StringComparison.Ordinal)) {
            content = BuildSyntheticFreeTextAppearance(objects, annotation, x, y, width, height, out fontResources);
        } else if (string.Equals(subtype, "Highlight", StringComparison.Ordinal)) {
            content = BuildSyntheticHighlightAppearance(objects, annotation, x, y, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (IsTextMarkupSubtype(subtype)) {
            content = BuildSyntheticTextMarkupAppearance(objects, annotation, subtype, x, y, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (IsShapeSubtype(subtype)) {
            content = BuildSyntheticShapeAppearance(objects, annotation, subtype, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (string.Equals(subtype, "Line", StringComparison.Ordinal)) {
            content = BuildSyntheticLineAppearance(objects, annotation, x, y, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (string.Equals(subtype, "Ink", StringComparison.Ordinal)) {
            content = BuildSyntheticInkAppearance(objects, annotation, x, y, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (IsPathAnnotationSubtype(subtype)) {
            content = BuildSyntheticPathAnnotationAppearance(objects, annotation, subtype, x, y, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else if (string.Equals(subtype, "Stamp", StringComparison.Ordinal)) {
            content = BuildSyntheticStampAppearance(objects, annotation, width, height);
            fontResources = new[] { ("Helv", PdfStandardFont.Helvetica) };
        } else if (string.Equals(subtype, "Caret", StringComparison.Ordinal)) {
            content = BuildSyntheticCaretAppearance(objects, annotation, width, height);
            fontResources = Array.Empty<(string Name, PdfStandardFont Font)>();
        } else {
            throw new NotSupportedException(UnsupportedVisualAnnotationMessage);
        }

        bool usesHighlightBlendMode = string.Equals(subtype, "Highlight", StringComparison.Ordinal);
        if (opacity.HasValue && !usesHighlightBlendMode) {
            content = ApplyAnnotationOpacity(content);
        }

        int appearanceObjectNumber = nextObjectNumber++;
        objects[appearanceObjectNumber] = new PdfIndirectObject(
            appearanceObjectNumber,
            0,
            CreateSyntheticAppearanceStream(width, height, content, fontResources, usesHighlightBlendMode, opacity));
        return new PdfReference(appearanceObjectNumber, 0);
    }

    private static string BuildSyntheticFreeTextAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height, out IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources) {
        string contents = TryReadFreeTextContents(objects, annotation);
        PdfFreeTextDefaultStyle defaultStyle = PdfFreeTextStyleParser.ParseDefaultStyle(TryReadString(objects, annotation, "DS"));
        double fontSize = TryReadFreeTextFontSize(objects, annotation) ?? defaultStyle.FontSize ?? 10D;
        PdfColor? textColor = TryReadFreeTextTextColor(objects, annotation) ?? defaultStyle.TextColor;
        PdfColor? borderColor = TryReadColor(objects, annotation, "C");
        double borderWidth = TryReadBorderWidth(objects, annotation);
        IReadOnlyList<double>? borderDashPattern = TryReadBorderDashPattern(objects, annotation);
        PdfFormFieldBorderStyle borderStyle = TryReadBorderStyle(objects, annotation);
        PdfColor? fillColor = TryReadColor(objects, annotation, "IC");
        PdfAlign textAlign = TryReadFreeTextAlignment(objects, annotation, defaultStyle.TextAlign);
        IReadOnlyList<PdfAnnotationPathPoint> calloutLine = TryReadFreeTextCalloutLine(objects, annotation, x, y, width, height);
        string? calloutLineEnding = TryReadFreeTextCalloutLineEnding(objects, annotation);
        double[]? rectangleDifferences = TryReadFreeTextRectangleDifferences(objects, annotation, width, height);
        TryReadBorderEffect(objects, annotation, out string? borderEffectStyle, out double borderEffectIntensity);
        IReadOnlyList<PdfFreeTextRichTextRun>? richRuns = PdfFreeTextStyleParser.ExtractRichTextRuns(TryReadString(objects, annotation, "RC"));
        if (richRuns != null) {
            return PdfAnnotationDictionaryBuilder.BuildFreeTextRichAppearanceContent(
                width,
                height,
                richRuns,
                out fontResources,
                fontSize,
                textColor,
                borderColor,
                borderWidth,
                fillColor,
                textAlign,
                borderDashPattern: borderDashPattern,
                borderStyle: borderStyle,
                calloutLine: calloutLine,
                calloutLineEnding: calloutLineEnding,
                rectangleDifferences: rectangleDifferences,
                borderEffectStyle: borderEffectStyle,
                borderEffectIntensity: borderEffectIntensity);
        }

        fontResources = new[] { ("Helv", PdfStandardFont.Helvetica) };
        return PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(
            width,
            height,
            contents,
            fontSize,
            textColor,
            borderColor,
            borderWidth,
            fillColor,
            textAlign,
            borderDashPattern: borderDashPattern,
            borderStyle: borderStyle,
            calloutLine: calloutLine,
            calloutLineEnding: calloutLineEnding,
            rectangleDifferences: rectangleDifferences,
            borderEffectStyle: borderEffectStyle,
            borderEffectIntensity: borderEffectIntensity);
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
        TryReadBorderEffect(objects, annotation, out string? borderEffectStyle, out double borderEffectIntensity);
        return PdfAnnotationDictionaryBuilder.BuildShapeAppearanceContent(
            width,
            height,
            subtype,
            TryReadColor(objects, annotation, "C"),
            TryReadColor(objects, annotation, "IC"),
            TryReadBorderWidth(objects, annotation),
            TryReadBorderDashPattern(objects, annotation),
            borderEffectStyle,
            borderEffectIntensity);
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
            TryReadBorderWidth(objects, annotation),
            TryReadBorderDashPattern(objects, annotation));
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
            TryReadBorderDashPattern(objects, annotation),
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
            TryReadBorderWidth(objects, annotation),
            TryReadBorderDashPattern(objects, annotation));
    }

    private static string BuildSyntheticCaretAppearance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double width, double height) {
        return PdfAnnotationDictionaryBuilder.BuildCaretAppearanceContent(
            width,
            height,
            TryReadColor(objects, annotation, "C"),
            TryReadBorderWidth(objects, annotation),
            TryReadBorderDashPattern(objects, annotation));
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
            TryReadBorderDashPattern(objects, annotation),
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

    private static IReadOnlyList<PdfAnnotationPathPoint> TryReadFreeTextCalloutLine(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double x, double y, double width, double height) {
        if (!annotation.Items.TryGetValue("CL", out var calloutObject) ||
            ResolveObject(objects, calloutObject) is not PdfArray calloutLine ||
            (calloutLine.Items.Count != 4 && calloutLine.Items.Count != 6)) {
            return Array.Empty<PdfAnnotationPathPoint>();
        }

        return ReadPathPoints(objects, calloutLine, x, y, width, height);
    }

    private static string? TryReadFreeTextCalloutLineEnding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (!annotation.Items.TryGetValue("LE", out var lineEndingObject)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(objects, lineEndingObject);
        if (resolved is PdfName lineEndingName) {
            return lineEndingName.Name;
        }

        if (resolved is PdfArray lineEndings &&
            lineEndings.Items.Count > 0 &&
            ResolveObject(objects, lineEndings.Items[0]) is PdfName firstEndingName) {
            return firstEndingName.Name;
        }

        return null;
    }

    private static double[]? TryReadFreeTextRectangleDifferences(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, double width, double height) {
        if (!annotation.Items.TryGetValue("RD", out var rectangleDifferencesObject) ||
            ResolveObject(objects, rectangleDifferencesObject) is not PdfArray rectangleDifferences ||
            rectangleDifferences.Items.Count < 4 ||
            !TryReadNonNegativeFiniteNumber(objects, rectangleDifferences.Items[0], out double left) ||
            !TryReadNonNegativeFiniteNumber(objects, rectangleDifferences.Items[1], out double top) ||
            !TryReadNonNegativeFiniteNumber(objects, rectangleDifferences.Items[2], out double right) ||
            !TryReadNonNegativeFiniteNumber(objects, rectangleDifferences.Items[3], out double bottom) ||
            left + right >= width ||
            top + bottom >= height) {
            return null;
        }

        return new[] { left, top, right, bottom };
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

    private static string ApplyAnnotationOpacity(string content) {
        return "q\n/OfficeIMOAnnotationOpacityGs gs\n" + content + "Q\n";
    }

    private static PdfStream CreateSyntheticAppearanceStream(double width, double height, string content, IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources, bool usesHighlightBlendMode, double? opacity) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        var bbox = new PdfArray();
        bbox.Items.Add(new PdfNumber(0D));
        bbox.Items.Add(new PdfNumber(0D));
        bbox.Items.Add(new PdfNumber(width));
        bbox.Items.Add(new PdfNumber(height));
        dictionary.Items["BBox"] = bbox;

        if (fontResources.Count > 0 || usesHighlightBlendMode || opacity.HasValue) {
            dictionary.Items["Resources"] = CreateSyntheticAppearanceResources(fontResources, usesHighlightBlendMode, opacity);
        }

        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfDictionary CreateSyntheticAppearanceResources(IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources, bool usesHighlightBlendMode, double? opacity) {
        var resources = new PdfDictionary();
        if (fontResources.Count > 0) {
            var fonts = new PdfDictionary();
            for (int i = 0; i < fontResources.Count; i++) {
                (string name, PdfStandardFont font) = fontResources[i];
                fonts.Items[name] = PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(font);
            }

            resources.Items["Font"] = fonts;
        }

        if (usesHighlightBlendMode || opacity.HasValue) {
            var extGStates = new PdfDictionary();
            if (usesHighlightBlendMode) {
                double highlightOpacity = opacity ?? 0.35D;
                var highlightState = new PdfDictionary();
                highlightState.Items["Type"] = new PdfName("ExtGState");
                highlightState.Items["BM"] = new PdfName("Multiply");
                highlightState.Items["CA"] = new PdfNumber(highlightOpacity);
                highlightState.Items["ca"] = new PdfNumber(highlightOpacity);
                extGStates.Items["OfficeIMOHighlightGs"] = highlightState;
            }

            if (opacity.HasValue && !usesHighlightBlendMode) {
                var opacityState = new PdfDictionary();
                opacityState.Items["Type"] = new PdfName("ExtGState");
                opacityState.Items["CA"] = new PdfNumber(opacity.Value);
                opacityState.Items["ca"] = new PdfNumber(opacity.Value);
                extGStates.Items["OfficeIMOAnnotationOpacityGs"] = opacityState;
            }

            resources.Items["ExtGState"] = extGStates;
        }

        return resources;
    }

    private static double? TryReadAnnotationOpacity(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        double? value = TryReadNumber(objects, annotation, "CA");
        if (!value.HasValue ||
            double.IsNaN(value.Value) ||
            double.IsInfinity(value.Value) ||
            value.Value < 0D ||
            value.Value > 1D) {
            return null;
        }

        return value.Value;
    }

    private static double TryReadBorderWidth(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is PdfDictionary borderStyle &&
            borderStyle.Items.TryGetValue("W", out PdfObject? borderStyleWidthObject) &&
            TryReadNonNegativeFiniteNumber(objects, borderStyleWidthObject, out double borderStyleWidth)) {
            return borderStyleWidth;
        }

        if (!annotation.Items.TryGetValue("Border", out var borderObject) ||
            ResolveObject(objects, borderObject) is not PdfArray border ||
            border.Items.Count < 3 ||
            !TryReadNonNegativeFiniteNumber(objects, border.Items[2], out double width)) {
            return 1D;
        }

        return width;
    }

    private static bool TryReadNonNegativeFiniteNumber(Dictionary<int, PdfIndirectObject> objects, PdfObject numberObject, out double value) {
        value = 0D;
        if (ResolveObject(objects, numberObject) is not PdfNumber number ||
            number.Value < 0D ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value)) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private static double[]? TryReadBorderDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyle ||
            borderStyle.Get<PdfName>("S")?.Name != "D") {
            return null;
        }

        if (!borderStyle.Items.TryGetValue("D", out PdfObject? dashObject)) {
            return new[] { 3D };
        }

        return TryReadDashPattern(objects, dashObject);
    }

    private static double[]? TryReadDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfObject dashObject) {
        if (ResolveObject(objects, dashObject) is not PdfArray dashArray || dashArray.Items.Count == 0) {
            return null;
        }

        var values = new double[dashArray.Items.Count];
        bool hasPositiveSegment = false;
        for (int i = 0; i < dashArray.Items.Count; i++) {
            if (!TryReadNonNegativeFiniteNumber(objects, dashArray.Items[i], out double segment)) {
                return null;
            }

            if (segment > 0D) {
                hasPositiveSegment = true;
            }

            values[i] = segment;
        }

        return hasPositiveSegment ? values : null;
    }

    private static PdfFormFieldBorderStyle TryReadBorderStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyle ||
            borderStyle.Get<PdfName>("S") is not PdfName styleName) {
            return PdfFormFieldBorderStyle.Solid;
        }

        switch (styleName.Name) {
            case "D":
                return PdfFormFieldBorderStyle.Dashed;
            case "U":
                return PdfFormFieldBorderStyle.Underline;
            case "B":
                return PdfFormFieldBorderStyle.Beveled;
            case "I":
                return PdfFormFieldBorderStyle.Inset;
            default:
                return PdfFormFieldBorderStyle.Solid;
        }
    }

    private static void TryReadBorderEffect(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, out string? style, out double intensity) {
        style = null;
        intensity = 0D;
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("BE", out PdfObject? borderEffectObject) ? borderEffectObject : null) is not PdfDictionary borderEffect ||
            borderEffect.Get<PdfName>("S") is not PdfName styleName ||
            !string.Equals(styleName.Name, "C", StringComparison.Ordinal)) {
            return;
        }

        double rawIntensity = TryReadNumber(objects, borderEffect, "I") ?? 1D;
        if (double.IsNaN(rawIntensity) || double.IsInfinity(rawIntensity) || rawIntensity <= 0D) {
            return;
        }

        style = styleName.Name;
        intensity = Math.Min(2D, rawIntensity);
    }

    private static string TryReadFreeTextContents(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        string? contents = TryReadString(objects, annotation, "Contents");
        if (!string.IsNullOrWhiteSpace(contents)) {
            return contents!;
        }

        return PdfFreeTextStyleParser.ExtractPlainText(TryReadString(objects, annotation, "RC")) ?? string.Empty;
    }

    private static PdfAlign TryReadFreeTextAlignment(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, PdfAlign? defaultAlignment) {
        double? alignment = TryReadNumber(objects, annotation, "Q");
        if (!alignment.HasValue) {
            return defaultAlignment ?? PdfAlign.Left;
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
        return PdfDefaultAppearanceParser.TryReadFontSize(defaultAppearance, out double fontSize)
            ? fontSize
            : null;
    }

    private static PdfColor? TryReadFreeTextTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        string? defaultAppearance = TryReadString(objects, annotation, "DA");
        return PdfDefaultAppearanceParser.TryReadTextColor(defaultAppearance, out PdfColor color)
            ? color
            : null;
    }
}
