namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationDictionaryBuilder {
    private static readonly char[] SpaceSeparators = { ' ' };
    private static readonly double[] DefaultBorderDashPattern = { 3D };

    internal static string BuildTextAnnotation(double x1, double y1, double x2, double y2, string contents, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, PdfColor? color = null, bool open = false, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        PdfDocument.ValidateTextAnnotationIcon(icon, nameof(icon));

        return "<< /Type /Annot /Subtype /Text /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /Contents " +
            PdfSyntaxEscaper.TextString(contents) +
            " /Name /" +
            PdfSyntaxEscaper.Name(GetTextAnnotationIconName(icon)) +
            (color.HasValue ? " /C [" + FormatCoordinate(color.Value.R) + " " + FormatCoordinate(color.Value.G) + " " + FormatCoordinate(color.Value.B) + "]" : string.Empty) +
            (open ? " /Open true" : string.Empty) +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildFreeTextAnnotation(double x1, double y1, double x2, double y2, string contents, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, int normalAppearanceId = 0, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        ValidateFinite(fontSize, nameof(fontSize));
        if (fontSize <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "PDF free text annotation font size must be a positive finite number.");
        }

        ValidateFinite(borderWidth, nameof(borderWidth));
        if (borderWidth < 0D) {
            throw new ArgumentOutOfRangeException(nameof(borderWidth), borderWidth, "PDF free text annotation border width must be non-negative.");
        }

        PdfColor resolvedTextColor = textColor ?? PdfColor.Black;
        return "<< /Type /Annot /Subtype /FreeText /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /Contents " +
            PdfSyntaxEscaper.TextString(contents) +
            " /DA " +
            PdfSyntaxEscaper.LiteralString("/Helv " + FormatCoordinate(fontSize) + " Tf " + FormatColor(resolvedTextColor) + " rg") +
            " /Border [0 0 " +
            FormatCoordinate(borderWidth) +
            "]" +
            (borderColor.HasValue && borderWidth > 0D ? " /C [" + FormatColor(borderColor.Value) + "]" : string.Empty) +
            (fillColor.HasValue ? " /IC [" + FormatColor(fillColor.Value) + "]" : string.Empty) +
            BuildNormalAppearanceEntry(normalAppearanceId) +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildHighlightAnnotation(double x1, double y1, double x2, double y2, string contents, PdfColor? color = null, int normalAppearanceId = 0, int? structParentIndex = null) {
        ValidateRectangle(x1, y1, x2, y2);
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        PdfColor resolvedColor = color ?? new PdfColor(1D, 0.92D, 0.2D);

        return "<< /Type /Annot /Subtype /Highlight /Rect [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) +
            "] /Contents " +
            PdfSyntaxEscaper.TextString(contents) +
            " /C [" +
            FormatColor(resolvedColor) +
            "] /QuadPoints [" +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y2) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y2) + " " +
            FormatCoordinate(x1) + " " +
            FormatCoordinate(y1) + " " +
            FormatCoordinate(x2) + " " +
            FormatCoordinate(y1) +
            "]" +
            BuildNormalAppearanceEntry(normalAppearanceId) +
            BuildStructParentEntry(structParentIndex) +
            " >>\n";
    }

    internal static string BuildFreeTextAppearanceContent(double width, double height, string contents, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null, IReadOnlyList<double>? borderDashPattern = null, PdfFormFieldBorderStyle borderStyle = PdfFormFieldBorderStyle.Solid, IReadOnlyList<PdfAnnotationPathPoint>? calloutLine = null, string? calloutLineEnding = null, IReadOnlyList<double>? rectangleDifferences = null, string? borderEffectStyle = null, double borderEffectIntensity = 0D) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(contents, nameof(contents));
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        Guard.LeftCenterRightAlign(textAlign, nameof(textAlign), "PDF free text annotation text");
        Guard.NonNegative(padding, nameof(padding));
        if (lineHeight.HasValue) {
            Guard.Positive(lineHeight.Value, nameof(lineHeight));
        }

        PdfColor resolvedTextColor = textColor ?? PdfColor.Black;
        ResolveFreeTextInnerRectangle(width, height, rectangleDifferences, out double innerX, out double innerY, out double innerWidth, out double innerHeight);
        double effectiveLineHeight = lineHeight ?? fontSize * 1.2D;
        double availableWidth = Math.Max(0D, innerWidth - padding * 2D);
        double availableHeight = Math.Max(0D, innerHeight - padding * 2D);
        List<string> lines = BuildFreeTextAppearanceLines(contents, fontSize, availableWidth);
        int maxVisibleLines = availableHeight > 0D ? Math.Max(1, (int)Math.Floor(availableHeight / effectiveLineHeight)) : 0;
        if (maxVisibleLines > 0 && lines.Count > maxVisibleLines) {
            var visibleLines = new List<string>(maxVisibleLines);
            for (int i = 0; i < maxVisibleLines; i++) {
                visibleLines.Add(lines[i]);
            }

            lines = visibleLines;
        }

        string content = "q\n";
        if (fillColor.HasValue) {
            content += FormatColor(fillColor.Value) + " rg " + FormatCoordinate(innerX) + " " + FormatCoordinate(innerY) + " " + FormatCoordinate(innerWidth) + " " + FormatCoordinate(innerHeight) + " re f\n";
        }

        if (calloutLine != null && calloutLine.Count >= 2) {
            PdfColor calloutColor = borderColor ?? resolvedTextColor;
            double calloutWidth = borderWidth > 0D ? borderWidth : 1D;
            content += BuildFreeTextCalloutLineContent(calloutLine, calloutColor, calloutWidth, borderDashPattern, calloutLineEnding);
        }

        if (borderColor.HasValue && borderWidth > 0D) {
            IReadOnlyList<double>? effectiveDashPattern = borderStyle == PdfFormFieldBorderStyle.Dashed && (borderDashPattern == null || borderDashPattern.Count == 0)
                ? DefaultBorderDashPattern
                : borderDashPattern;
            string borderContent = string.Equals(borderEffectStyle, "C", StringComparison.Ordinal) && borderEffectIntensity > 0D
                ? BuildShapeAppearanceContent(innerWidth, innerHeight, "Square", borderColor.Value, null, borderWidth, effectiveDashPattern, borderEffectStyle, borderEffectIntensity)
                : PdfAcroFormDictionaryBuilder.BuildRectangularBorderAppearanceContent(innerWidth, innerHeight, borderColor.Value, borderWidth, effectiveDashPattern, borderStyle);
            content += TranslateContent(borderContent, innerX, innerY);
        }

        double baseline = innerY + innerHeight - padding - fontSize;
        double minimumBaseline = innerY + padding - fontSize * 0.25D;
        for (int i = 0; i < lines.Count && baseline >= minimumBaseline; i++) {
            string line = lines[i];
            if (!PdfWinAnsiEncoding.CanEncode(line, out _)) {
                baseline -= effectiveLineHeight;
                continue;
            }

            double lineWidth = EstimateWinAnsiTextWidth(line, fontSize);
            double textX = innerX + ResolveFreeTextLineX(textAlign, padding, availableWidth, lineWidth);
            content += "BT /Helv " + FormatCoordinate(fontSize) + " Tf " + FormatColor(resolvedTextColor) + " rg " + FormatCoordinate(textX) + " " + FormatCoordinate(Math.Max(0D, baseline)) + " Td " + PdfSyntaxEscaper.WinAnsiHexString(line) + " Tj ET\n";
            baseline -= effectiveLineHeight;
        }

        return content + "Q\n";
    }

    internal static string BuildFreeTextRichAppearanceContent(double width, double height, IReadOnlyList<PdfFreeTextRichTextRun> runs, out IReadOnlyList<(string Name, PdfStandardFont Font)> fontResources, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null, IReadOnlyList<double>? borderDashPattern = null, PdfFormFieldBorderStyle borderStyle = PdfFormFieldBorderStyle.Solid, IReadOnlyList<PdfAnnotationPathPoint>? calloutLine = null, string? calloutLineEnding = null, IReadOnlyList<double>? rectangleDifferences = null, string? borderEffectStyle = null, double borderEffectIntensity = 0D) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(runs, nameof(runs));
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        Guard.LeftCenterRightAlign(textAlign, nameof(textAlign), "PDF free text annotation text");
        Guard.NonNegative(padding, nameof(padding));
        if (lineHeight.HasValue) {
            Guard.Positive(lineHeight.Value, nameof(lineHeight));
        }

        PdfColor resolvedTextColor = textColor ?? PdfColor.Black;
        ResolveFreeTextInnerRectangle(width, height, rectangleDifferences, out double innerX, out double innerY, out double innerWidth, out double innerHeight);
        double effectiveLineHeight = lineHeight ?? fontSize * 1.2D;
        double availableWidth = Math.Max(0D, innerWidth - padding * 2D);
        double availableHeight = Math.Max(0D, innerHeight - padding * 2D);
        List<List<PdfFreeTextRichTextRun>> lines = BuildFreeTextRichAppearanceLines(runs, fontSize, availableWidth);
        int maxVisibleLines = availableHeight > 0D ? Math.Max(1, (int)Math.Floor(availableHeight / effectiveLineHeight)) : 0;
        if (maxVisibleLines > 0 && lines.Count > maxVisibleLines) {
            lines = lines.GetRange(0, maxVisibleLines);
        }

        var resources = new List<(string Name, PdfStandardFont Font)>();
        var content = new StringBuilder();
        content.Append("q\n");
        if (fillColor.HasValue) {
            content.Append(FormatColor(fillColor.Value)).Append(" rg ")
                .Append(FormatCoordinate(innerX)).Append(' ')
                .Append(FormatCoordinate(innerY)).Append(' ')
                .Append(FormatCoordinate(innerWidth)).Append(' ')
                .Append(FormatCoordinate(innerHeight)).Append(" re f\n");
        }

        if (calloutLine != null && calloutLine.Count >= 2) {
            PdfColor calloutColor = borderColor ?? resolvedTextColor;
            double calloutWidth = borderWidth > 0D ? borderWidth : 1D;
            content.Append(BuildFreeTextCalloutLineContent(calloutLine, calloutColor, calloutWidth, borderDashPattern, calloutLineEnding));
        }

        if (borderColor.HasValue && borderWidth > 0D) {
            IReadOnlyList<double>? effectiveDashPattern = borderStyle == PdfFormFieldBorderStyle.Dashed && (borderDashPattern == null || borderDashPattern.Count == 0)
                ? DefaultBorderDashPattern
                : borderDashPattern;
            string borderContent = string.Equals(borderEffectStyle, "C", StringComparison.Ordinal) && borderEffectIntensity > 0D
                ? BuildShapeAppearanceContent(innerWidth, innerHeight, "Square", borderColor.Value, null, borderWidth, effectiveDashPattern, borderEffectStyle, borderEffectIntensity)
                : PdfAcroFormDictionaryBuilder.BuildRectangularBorderAppearanceContent(innerWidth, innerHeight, borderColor.Value, borderWidth, effectiveDashPattern, borderStyle);
            content.Append(TranslateContent(borderContent, innerX, innerY));
        }

        double baseline = innerY + innerHeight - padding - fontSize;
        double minimumBaseline = innerY + padding - fontSize * 0.25D;
        for (int i = 0; i < lines.Count && baseline >= minimumBaseline; i++) {
            List<PdfFreeTextRichTextRun> line = lines[i];
            if (!CanEncodeFreeTextRichLine(line)) {
                baseline -= effectiveLineHeight;
                continue;
            }

            double lineWidth = MeasureFreeTextRichLine(line, fontSize);
            double textX = innerX + ResolveFreeTextLineX(textAlign, padding, availableWidth, lineWidth);
            AppendFreeTextRichLine(content, line, textX, Math.Max(0D, baseline), fontSize, resolvedTextColor, resources);
            baseline -= effectiveLineHeight;
        }

        content.Append("Q\n");
        fontResources = resources.Count == 0
            ? new[] { ("Helv", PdfStandardFont.Helvetica) }
            : resources;
        return content.ToString();
    }

    private static void ResolveFreeTextInnerRectangle(double width, double height, IReadOnlyList<double>? rectangleDifferences, out double x, out double y, out double innerWidth, out double innerHeight) {
        x = 0D;
        y = 0D;
        innerWidth = width;
        innerHeight = height;
        if (rectangleDifferences == null || rectangleDifferences.Count < 4) {
            return;
        }

        double left = rectangleDifferences[0];
        double top = rectangleDifferences[1];
        double right = rectangleDifferences[2];
        double bottom = rectangleDifferences[3];
        if (!IsValidRectangleDifference(left) ||
            !IsValidRectangleDifference(top) ||
            !IsValidRectangleDifference(right) ||
            !IsValidRectangleDifference(bottom) ||
            left + right >= width ||
            top + bottom >= height) {
            return;
        }

        x = left;
        y = bottom;
        innerWidth = width - left - right;
        innerHeight = height - top - bottom;
    }

    private static bool IsValidRectangleDifference(double value) =>
        value >= 0D && !double.IsNaN(value) && !double.IsInfinity(value);

    private static string TranslateContent(string content, double x, double y) {
        if (string.IsNullOrEmpty(content) || (Math.Abs(x) <= 0.0001D && Math.Abs(y) <= 0.0001D)) {
            return content;
        }

        return "q\n1 0 0 1 " + FormatCoordinate(x) + " " + FormatCoordinate(y) + " cm\n" + content + "Q\n";
    }

    private static string BuildFreeTextCalloutLineContent(IReadOnlyList<PdfAnnotationPathPoint> calloutLine, PdfColor color, double borderWidth, IReadOnlyList<double>? borderDashPattern, string? lineEnding) {
        var builder = new StringBuilder();
        builder.Append(BuildBorderStrokeOperators(color, borderWidth, borderDashPattern));
        builder.Append(FormatCoordinate(calloutLine[0].X)).Append(' ').Append(FormatCoordinate(calloutLine[0].Y)).Append(" m ");
        for (int i = 1; i < calloutLine.Count; i++) {
            builder.Append(FormatCoordinate(calloutLine[i].X)).Append(' ').Append(FormatCoordinate(calloutLine[i].Y)).Append(" l ");
        }

        builder.Append("S\n");
        AppendLineEnding(builder, lineEnding, calloutLine[0].X, calloutLine[0].Y, calloutLine[1].X, calloutLine[1].Y, borderWidth, color, color);
        return builder.ToString();
    }

    internal static string BuildHighlightAppearanceContent(double width, double height, PdfColor? color = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        PdfColor resolvedColor = color ?? new PdfColor(1D, 0.92D, 0.2D);
        return "q\n" +
            "/OfficeIMOHighlightGs gs\n" +
            FormatColor(resolvedColor) + " rg 0 0 " + FormatCoordinate(width) + " " + FormatCoordinate(height) + " re f\n" +
            "Q\n";
    }

    internal static string BuildHighlightAppearanceContent(double width, double height, IReadOnlyList<PdfHighlightQuad> quadPoints, PdfColor? color = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(quadPoints, nameof(quadPoints));
        if (quadPoints.Count == 0) {
            return BuildHighlightAppearanceContent(width, height, color);
        }

        PdfColor resolvedColor = color ?? new PdfColor(1D, 0.92D, 0.2D);
        var builder = new StringBuilder();
        builder.Append("q\n")
            .Append("/OfficeIMOHighlightGs gs\n")
            .Append(FormatColor(resolvedColor))
            .Append(" rg\n");

        for (int i = 0; i < quadPoints.Count; i++) {
            PdfHighlightQuad quad = quadPoints[i];
            builder.Append(FormatCoordinate(quad.X1)).Append(' ').Append(FormatCoordinate(quad.Y1)).Append(" m ")
                .Append(FormatCoordinate(quad.X2)).Append(' ').Append(FormatCoordinate(quad.Y2)).Append(" l ")
                .Append(FormatCoordinate(quad.X4)).Append(' ').Append(FormatCoordinate(quad.Y4)).Append(" l ")
                .Append(FormatCoordinate(quad.X3)).Append(' ').Append(FormatCoordinate(quad.Y3)).Append(" l h f\n");
        }

        builder.Append("Q\n");
        return builder.ToString();
    }

    internal static string BuildTextMarkupAppearanceContent(double width, double height, string subtype, PdfColor? color = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateTextMarkupSubtype(subtype);
        PdfColor resolvedColor = color ?? PdfColor.Black;
        double y = ResolveTextMarkupLineY(subtype, 0D, height);
        if (string.Equals(subtype, "Squiggly", StringComparison.Ordinal)) {
            var builder = new StringBuilder();
            builder.Append("q\n")
                .Append(FormatColor(resolvedColor))
                .Append(" RG 1 w\n");
            AppendSquigglyLine(builder, 0D, width, y, ResolveSquigglyAmplitude(0D, height));
            builder.Append("Q\n");
            return builder.ToString();
        }

        return "q\n" +
            FormatColor(resolvedColor) + " RG 1 w 0 " + FormatCoordinate(y) + " m " + FormatCoordinate(width) + " " + FormatCoordinate(y) + " l S\n" +
            "Q\n";
    }

    internal static string BuildTextMarkupAppearanceContent(double width, double height, IReadOnlyList<PdfHighlightQuad> quadPoints, string subtype, PdfColor? color = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(quadPoints, nameof(quadPoints));
        ValidateTextMarkupSubtype(subtype);
        if (quadPoints.Count == 0) {
            return BuildTextMarkupAppearanceContent(width, height, subtype, color);
        }

        PdfColor resolvedColor = color ?? PdfColor.Black;
        var builder = new StringBuilder();
        builder.Append("q\n")
            .Append(FormatColor(resolvedColor))
            .Append(" RG 1 w\n");

        for (int i = 0; i < quadPoints.Count; i++) {
            PdfHighlightQuad quad = quadPoints[i];
            double startX = Math.Min(Math.Min(quad.X1, quad.X3), Math.Min(quad.X2, quad.X4));
            double endX = Math.Max(Math.Max(quad.X1, quad.X3), Math.Max(quad.X2, quad.X4));
            double bottomY = Math.Min(Math.Min(quad.Y1, quad.Y2), Math.Min(quad.Y3, quad.Y4));
            double topY = Math.Max(Math.Max(quad.Y1, quad.Y2), Math.Max(quad.Y3, quad.Y4));
            double lineY = ResolveTextMarkupLineY(subtype, bottomY, topY);
            if (string.Equals(subtype, "Squiggly", StringComparison.Ordinal)) {
                AppendSquigglyLine(builder, startX, endX, lineY, ResolveSquigglyAmplitude(bottomY, topY));
            } else {
                builder.Append(FormatCoordinate(startX)).Append(' ').Append(FormatCoordinate(lineY)).Append(" m ")
                    .Append(FormatCoordinate(endX)).Append(' ').Append(FormatCoordinate(lineY)).Append(" l S\n");
            }
        }

        builder.Append("Q\n");
        return builder.ToString();
    }

    internal static string BuildShapeAppearanceContent(double width, double height, string subtype, PdfColor? strokeColor = null, PdfColor? fillColor = null, double borderWidth = 1D, IReadOnlyList<double>? borderDashPattern = null, string? borderEffectStyle = null, double borderEffectIntensity = 0D) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        Guard.NonNegative(borderEffectIntensity, nameof(borderEffectIntensity));
        ValidateShapeSubtype(subtype);

        bool hasStroke = borderWidth > 0D;
        if (!hasStroke && !fillColor.HasValue) {
            return "q\nQ\n";
        }

        PdfColor resolvedStrokeColor = strokeColor ?? PdfColor.Black;
        double inset = hasStroke ? borderWidth * 0.5D : 0D;
        bool useCloudyBorder = hasStroke &&
            borderEffectIntensity > 0D &&
            string.Equals(borderEffectStyle, "C", StringComparison.Ordinal);
        var builder = new StringBuilder();
        builder.Append("q\n");
        if (fillColor.HasValue && useCloudyBorder) {
            builder.Append(FormatColor(fillColor.Value)).Append(" rg ");
            if (string.Equals(subtype, "Square", StringComparison.Ordinal)) {
                AppendSquarePath(builder, inset, inset, Math.Max(0D, width - inset * 2D), Math.Max(0D, height - inset * 2D));
            } else {
                AppendCirclePath(builder, inset, inset, Math.Max(0D, width - inset * 2D), Math.Max(0D, height - inset * 2D));
            }

            builder.Append("f\n");
        } else if (fillColor.HasValue) {
            builder.Append(FormatColor(fillColor.Value)).Append(" rg ");
        }

        if (hasStroke) {
            builder.Append(BuildBorderStrokeOperators(resolvedStrokeColor, borderWidth, borderDashPattern));
        }

        if (useCloudyBorder) {
            double cloudyInset = inset + ResolveCloudyBorderAmplitude(borderWidth, borderEffectIntensity);
            if (string.Equals(subtype, "Square", StringComparison.Ordinal)) {
                AppendCloudyRectanglePath(builder, cloudyInset, cloudyInset, Math.Max(0D, width - cloudyInset * 2D), Math.Max(0D, height - cloudyInset * 2D), borderWidth, borderEffectIntensity);
            } else {
                AppendCloudyCirclePath(builder, cloudyInset, cloudyInset, Math.Max(0D, width - cloudyInset * 2D), Math.Max(0D, height - cloudyInset * 2D), borderWidth, borderEffectIntensity);
            }

            builder.Append("S\n");
        } else if (string.Equals(subtype, "Square", StringComparison.Ordinal)) {
            AppendSquarePath(builder, inset, inset, Math.Max(0D, width - inset * 2D), Math.Max(0D, height - inset * 2D));
            builder.Append(fillColor.HasValue && hasStroke ? "B\n" : fillColor.HasValue ? "f\n" : "S\n");
        } else {
            AppendCirclePath(builder, inset, inset, Math.Max(0D, width - inset * 2D), Math.Max(0D, height - inset * 2D));
            builder.Append(fillColor.HasValue && hasStroke ? "B\n" : fillColor.HasValue ? "f\n" : "S\n");
        }

        builder.Append("Q\n");
        return builder.ToString();
    }

    internal static string BuildLineAppearanceContent(double width, double height, double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, PdfColor? fillColor = null, double borderWidth = 1D, IReadOnlyList<double>? borderDashPattern = null, string? startEnding = null, string? endEnding = null) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateFinite(x2, nameof(x2));
        ValidateFinite(y2, nameof(y2));
        if (Math.Abs(x1 - x2) <= 0.0001D && Math.Abs(y1 - y2) <= 0.0001D) {
            throw new ArgumentOutOfRangeException(nameof(x2), x2, "PDF line annotation endpoints must not be identical.");
        }

        if (borderWidth <= 0D) {
            return "q\nQ\n";
        }

        PdfColor resolvedStrokeColor = strokeColor ?? PdfColor.Black;
        PdfColor resolvedFillColor = fillColor ?? resolvedStrokeColor;
        var builder = new StringBuilder();
        builder.Append("q\n")
            .Append(BuildBorderStrokeOperators(resolvedStrokeColor, borderWidth, borderDashPattern))
            .Append(FormatCoordinate(x1)).Append(' ').Append(FormatCoordinate(y1)).Append(" m ")
            .Append(FormatCoordinate(x2)).Append(' ').Append(FormatCoordinate(y2)).Append(" l S\n");

        AppendLineEnding(builder, startEnding, x1, y1, x2, y2, borderWidth, resolvedStrokeColor, resolvedFillColor);
        AppendLineEnding(builder, endEnding, x2, y2, x1, y1, borderWidth, resolvedStrokeColor, resolvedFillColor);
        builder.Append("Q\n");
        return builder.ToString();
    }

    private static List<string> BuildFreeTextAppearanceLines(string contents, double fontSize, double availableWidth) {
        string normalized = contents.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] paragraphs = normalized.Split('\n');
        var lines = new List<string>();
        for (int i = 0; i < paragraphs.Length; i++) {
            AddWrappedFreeTextParagraph(lines, paragraphs[i], fontSize, availableWidth);
        }

        if (lines.Count == 0) {
            lines.Add(string.Empty);
        }

        return lines;
    }

    private static List<List<PdfFreeTextRichTextRun>> BuildFreeTextRichAppearanceLines(IReadOnlyList<PdfFreeTextRichTextRun> runs, double defaultFontSize, double availableWidth) {
        var lines = new List<List<PdfFreeTextRichTextRun>>();
        var current = new List<PdfFreeTextRichTextRun>();
        double currentWidth = 0D;

        for (int i = 0; i < runs.Count; i++) {
            PdfFreeTextRichTextRun run = runs[i];
            if (run.IsLineBreak) {
                lines.Add(current);
                current = new List<PdfFreeTextRichTextRun>();
                currentWidth = 0D;
                continue;
            }

            AppendFreeTextRichWrappedRun(lines, current, ref currentWidth, run, defaultFontSize, availableWidth);
        }

        lines.Add(current);
        if (lines.Count == 0) {
            lines.Add(new List<PdfFreeTextRichTextRun>());
        }

        return lines;
    }

    private static void AppendFreeTextRichWrappedRun(List<List<PdfFreeTextRichTextRun>> lines, List<PdfFreeTextRichTextRun> current, ref double currentWidth, PdfFreeTextRichTextRun run, double defaultFontSize, double availableWidth) {
        string text = run.Text.Replace("\r\n", "\n").Replace('\r', '\n');
        int tokenStart = 0;
        bool tokenIsSpace = false;
        bool hasToken = false;
        for (int i = 0; i <= text.Length; i++) {
            bool atEnd = i == text.Length;
            bool isSpace = !atEnd && char.IsWhiteSpace(text[i]) && text[i] != '\n';
            bool isLineBreak = !atEnd && text[i] == '\n';
            if (!hasToken && !atEnd && !isLineBreak) {
                tokenStart = i;
                tokenIsSpace = isSpace;
                hasToken = true;
                continue;
            }

            if (hasToken && (atEnd || isLineBreak || isSpace != tokenIsSpace)) {
                string token = text.Substring(tokenStart, i - tokenStart);
                AppendFreeTextRichToken(lines, current, ref currentWidth, run, tokenIsSpace ? " " : token, tokenIsSpace, defaultFontSize, availableWidth);
                hasToken = false;
            }

            if (isLineBreak) {
                lines.Add(new List<PdfFreeTextRichTextRun>(current));
                current.Clear();
                currentWidth = 0D;
            }
        }
    }

    private static void AppendFreeTextRichToken(List<List<PdfFreeTextRichTextRun>> lines, List<PdfFreeTextRichTextRun> current, ref double currentWidth, PdfFreeTextRichTextRun template, string token, bool isWhitespace, double defaultFontSize, double availableWidth) {
        if (token.Length == 0 || (isWhitespace && current.Count == 0)) {
            return;
        }

        double tokenWidth = EstimateWinAnsiTextWidth(token, template.FontSize ?? defaultFontSize);
        if (!isWhitespace && availableWidth > 0D && current.Count > 0 && currentWidth + tokenWidth > availableWidth) {
            TrimTrailingWhitespace(current, ref currentWidth, defaultFontSize);
            lines.Add(new List<PdfFreeTextRichTextRun>(current));
            current.Clear();
            currentWidth = 0D;
        }

        PdfFreeTextRichTextRun segment = new PdfFreeTextRichTextRun(token, template.Bold, template.Italic, template.Underline, template.Strike, template.Color, template.FontSize, isLineBreak: false);
        current.Add(segment);
        currentWidth += tokenWidth;
    }

    private static void TrimTrailingWhitespace(List<PdfFreeTextRichTextRun> line, ref double width, double defaultFontSize) {
        while (line.Count > 0 && string.IsNullOrWhiteSpace(line[line.Count - 1].Text)) {
            PdfFreeTextRichTextRun run = line[line.Count - 1];
            width -= EstimateWinAnsiTextWidth(run.Text, run.FontSize ?? defaultFontSize);
            line.RemoveAt(line.Count - 1);
        }
    }

    private static bool CanEncodeFreeTextRichLine(List<PdfFreeTextRichTextRun> line) {
        for (int i = 0; i < line.Count; i++) {
            if (!PdfWinAnsiEncoding.CanEncode(line[i].Text, out _)) {
                return false;
            }
        }

        return true;
    }

    private static double MeasureFreeTextRichLine(List<PdfFreeTextRichTextRun> line, double defaultFontSize) {
        double width = 0D;
        for (int i = 0; i < line.Count; i++) {
            width += EstimateWinAnsiTextWidth(line[i].Text, line[i].FontSize ?? defaultFontSize);
        }

        return width;
    }

    private static void AppendFreeTextRichLine(StringBuilder content, List<PdfFreeTextRichTextRun> line, double x, double y, double defaultFontSize, PdfColor defaultTextColor, List<(string Name, PdfStandardFont Font)> resources) {
        double xCursor = 0D;
        for (int i = 0; i < line.Count; i++) {
            PdfFreeTextRichTextRun run = line[i];
            if (run.Text.Length == 0) {
                continue;
            }

            double runFontSize = run.FontSize ?? defaultFontSize;
            PdfColor runColor = run.Color ?? defaultTextColor;
            PdfStandardFont font = ResolveFreeTextRichFont(run);
            string resourceName = ResolveFreeTextRichFontResourceName(font);
            RegisterFreeTextRichFontResource(resources, resourceName, font);
            double segmentWidth = EstimateWinAnsiTextWidth(run.Text, runFontSize);
            content.Append("BT /")
                .Append(PdfSyntaxEscaper.Name(resourceName))
                .Append(' ')
                .Append(FormatCoordinate(runFontSize))
                .Append(" Tf ")
                .Append(FormatColor(runColor))
                .Append(" rg ")
                .Append(FormatCoordinate(x + xCursor))
                .Append(' ')
                .Append(FormatCoordinate(y))
                .Append(" Td ")
                .Append(PdfSyntaxEscaper.WinAnsiHexString(run.Text))
                .Append(" Tj ET\n");
            if (run.Underline || run.Strike) {
                AppendFreeTextRichDecoration(content, x + xCursor, y, segmentWidth, runFontSize, runColor, run.Underline, run.Strike);
            }

            xCursor += segmentWidth;
        }
    }

    private static void AppendFreeTextRichDecoration(StringBuilder content, double x, double y, double width, double fontSize, PdfColor color, bool underline, bool strike) {
        double strokeWidth = Math.Max(0.35D, fontSize / 18D);
        if (underline) {
            double underlineY = y - Math.Max(1D, fontSize * 0.12D);
            content.Append("q ").Append(FormatColor(color)).Append(" RG ")
                .Append(FormatCoordinate(strokeWidth)).Append(" w ")
                .Append(FormatCoordinate(x)).Append(' ')
                .Append(FormatCoordinate(underlineY)).Append(" m ")
                .Append(FormatCoordinate(x + width)).Append(' ')
                .Append(FormatCoordinate(underlineY)).Append(" l S Q\n");
        }

        if (strike) {
            double strikeY = y + fontSize * 0.32D;
            content.Append("q ").Append(FormatColor(color)).Append(" RG ")
                .Append(FormatCoordinate(strokeWidth)).Append(" w ")
                .Append(FormatCoordinate(x)).Append(' ')
                .Append(FormatCoordinate(strikeY)).Append(" m ")
                .Append(FormatCoordinate(x + width)).Append(' ')
                .Append(FormatCoordinate(strikeY)).Append(" l S Q\n");
        }
    }

    private static PdfStandardFont ResolveFreeTextRichFont(PdfFreeTextRichTextRun run) {
        if (run.Bold && run.Italic) {
            return PdfStandardFont.HelveticaBoldOblique;
        }

        if (run.Bold) {
            return PdfStandardFont.HelveticaBold;
        }

        return run.Italic ? PdfStandardFont.HelveticaOblique : PdfStandardFont.Helvetica;
    }

    private static string ResolveFreeTextRichFontResourceName(PdfStandardFont font) =>
        font switch {
            PdfStandardFont.HelveticaBold => "HelvB",
            PdfStandardFont.HelveticaOblique => "HelvI",
            PdfStandardFont.HelveticaBoldOblique => "HelvBI",
            _ => "Helv"
        };

    private static void RegisterFreeTextRichFontResource(List<(string Name, PdfStandardFont Font)> resources, string name, PdfStandardFont font) {
        for (int i = 0; i < resources.Count; i++) {
            if (string.Equals(resources[i].Name, name, StringComparison.Ordinal)) {
                return;
            }
        }

        resources.Add((name, font));
    }

    private static void AddWrappedFreeTextParagraph(List<string> lines, string paragraph, double fontSize, double availableWidth) {
        string trimmed = paragraph.Trim();
        if (trimmed.Length == 0) {
            lines.Add(string.Empty);
            return;
        }

        if (availableWidth <= 0D) {
            lines.Add(trimmed);
            return;
        }

        string[] words = trimmed.Split(SpaceSeparators, StringSplitOptions.RemoveEmptyEntries);
        var current = new StringBuilder();
        for (int i = 0; i < words.Length; i++) {
            string word = words[i];
            string candidate = current.Length == 0 ? word : current + " " + word;
            if (EstimateWinAnsiTextWidth(candidate, fontSize) <= availableWidth || current.Length == 0) {
                current.Clear();
                current.Append(candidate);
                continue;
            }

            lines.Add(current.ToString());
            current.Clear();
            current.Append(word);
        }

        if (current.Length > 0) {
            lines.Add(current.ToString());
        }
    }

    private static double ResolveFreeTextLineX(PdfAlign textAlign, double padding, double availableWidth, double lineWidth) {
        if (textAlign == PdfAlign.Center) {
            return padding + Math.Max(0D, (availableWidth - lineWidth) / 2D);
        }

        if (textAlign == PdfAlign.Right) {
            return padding + Math.Max(0D, availableWidth - lineWidth);
        }

        return padding;
    }

    private static double EstimateWinAnsiTextWidth(string text, double fontSize) {
        double units = 0D;
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch == ' ') {
                units += 250D;
            } else if (ch == 'i' || ch == 'l' || ch == 'I' || ch == '.' || ch == ',' || ch == '\'' || ch == ':' || ch == ';' || ch == '!') {
                units += 280D;
            } else if (ch == 'm' || ch == 'w' || ch == 'M' || ch == 'W') {
                units += 780D;
            } else {
                units += 500D;
            }
        }

        return units / 1000D * fontSize;
    }

    private static string GetTextAnnotationIconName(PdfTextAnnotationIcon icon) =>
        icon switch {
            PdfTextAnnotationIcon.Comment => "Comment",
            PdfTextAnnotationIcon.Key => "Key",
            PdfTextAnnotationIcon.Note => "Note",
            PdfTextAnnotationIcon.Help => "Help",
            PdfTextAnnotationIcon.NewParagraph => "NewParagraph",
            PdfTextAnnotationIcon.Paragraph => "Paragraph",
            PdfTextAnnotationIcon.Insert => "Insert",
            _ => throw new ArgumentOutOfRangeException(nameof(icon), "PDF text annotation icon must be Comment, Key, Note, Help, NewParagraph, Paragraph, or Insert.")
        };

    private static string BuildNormalAppearanceEntry(int normalAppearanceId) {
        if (normalAppearanceId == 0) {
            return string.Empty;
        }

        if (normalAppearanceId < 0) {
            throw new ArgumentOutOfRangeException(nameof(normalAppearanceId), normalAppearanceId, "PDF annotation appearance object id must be positive.");
        }

        return " /AP << /N " + PdfSyntaxEscaper.IndirectReference(normalAppearanceId) + " >>";
    }

    private static double ResolveTextMarkupLineY(string subtype, double bottomY, double topY) {
        if (string.Equals(subtype, "StrikeOut", StringComparison.Ordinal)) {
            return bottomY + (topY - bottomY) * 0.55D;
        }

        if (string.Equals(subtype, "Squiggly", StringComparison.Ordinal)) {
            return bottomY + ResolveSquigglyAmplitude(bottomY, topY);
        }

        return bottomY;
    }

    private static double ResolveSquigglyAmplitude(double bottomY, double topY) =>
        Math.Max(1D, Math.Min(2D, (topY - bottomY) * 0.18D));

    private static void AppendSquigglyLine(StringBuilder builder, double startX, double endX, double baseY, double amplitude) {
        const int maximumSegments = 512;
        double width = Math.Max(0D, endX - startX);
        double step = Math.Max(amplitude * 2D, width / maximumSegments);
        builder.Append(FormatCoordinate(startX)).Append(' ').Append(FormatCoordinate(baseY)).Append(" m ");
        double x = startX;
        bool up = true;
        int segmentCount = 0;
        while (x + step < endX && segmentCount++ < maximumSegments) {
            x += step;
            double y = up ? baseY + amplitude : baseY - amplitude;
            builder.Append(FormatCoordinate(x)).Append(' ').Append(FormatCoordinate(y)).Append(" l ");
            up = !up;
        }

        builder.Append(FormatCoordinate(endX)).Append(' ').Append(FormatCoordinate(baseY)).Append(" l S\n");
    }

    private static void ValidateTextMarkupSubtype(string subtype) {
        if (!string.Equals(subtype, "Underline", StringComparison.Ordinal) &&
            !string.Equals(subtype, "StrikeOut", StringComparison.Ordinal) &&
            !string.Equals(subtype, "Squiggly", StringComparison.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(subtype), subtype, "PDF text markup subtype must be Underline, StrikeOut, or Squiggly.");
        }
    }

    private static void AppendSquarePath(StringBuilder builder, double x, double y, double width, double height) {
        builder.Append(FormatCoordinate(x)).Append(' ')
            .Append(FormatCoordinate(y)).Append(' ')
            .Append(FormatCoordinate(width)).Append(' ')
            .Append(FormatCoordinate(height)).Append(" re ");
    }

    private static void AppendCirclePath(StringBuilder builder, double x, double y, double width, double height) {
        const double kappa = 0.552284749831D;
        double rx = width / 2D;
        double ry = height / 2D;
        double cx = x + rx;
        double cy = y + ry;
        double ox = rx * kappa;
        double oy = ry * kappa;

        builder.Append(FormatCoordinate(cx + rx)).Append(' ').Append(FormatCoordinate(cy)).Append(" m ")
            .Append(FormatCoordinate(cx + rx)).Append(' ').Append(FormatCoordinate(cy + oy)).Append(' ')
            .Append(FormatCoordinate(cx + ox)).Append(' ').Append(FormatCoordinate(cy + ry)).Append(' ')
            .Append(FormatCoordinate(cx)).Append(' ').Append(FormatCoordinate(cy + ry)).Append(" c ")
            .Append(FormatCoordinate(cx - ox)).Append(' ').Append(FormatCoordinate(cy + ry)).Append(' ')
            .Append(FormatCoordinate(cx - rx)).Append(' ').Append(FormatCoordinate(cy + oy)).Append(' ')
            .Append(FormatCoordinate(cx - rx)).Append(' ').Append(FormatCoordinate(cy)).Append(" c ")
            .Append(FormatCoordinate(cx - rx)).Append(' ').Append(FormatCoordinate(cy - oy)).Append(' ')
            .Append(FormatCoordinate(cx - ox)).Append(' ').Append(FormatCoordinate(cy - ry)).Append(' ')
            .Append(FormatCoordinate(cx)).Append(' ').Append(FormatCoordinate(cy - ry)).Append(" c ")
            .Append(FormatCoordinate(cx + ox)).Append(' ').Append(FormatCoordinate(cy - ry)).Append(' ')
            .Append(FormatCoordinate(cx + rx)).Append(' ').Append(FormatCoordinate(cy - oy)).Append(' ')
            .Append(FormatCoordinate(cx + rx)).Append(' ').Append(FormatCoordinate(cy)).Append(" c ");
    }

    private static double ResolveCloudyBorderAmplitude(double borderWidth, double intensity) =>
        Math.Max(borderWidth, 2D + intensity * 1.5D);

    private static double ResolveCloudyBorderStep(double borderWidth, double intensity) =>
        Math.Max(6D, borderWidth * 3D + intensity * 4D);

    private static void AppendCloudyRectanglePath(StringBuilder builder, double x, double y, double width, double height, double borderWidth, double intensity) {
        double amplitude = ResolveCloudyBorderAmplitude(borderWidth, intensity);
        double step = ResolveCloudyBorderStep(borderWidth, intensity);
        builder.Append(FormatCoordinate(x)).Append(' ').Append(FormatCoordinate(y)).Append(" m ");
        AppendCloudyEdge(builder, x, y, x + width, y, 0D, amplitude, step);
        AppendCloudyEdge(builder, x + width, y, x + width, y + height, -amplitude, 0D, step);
        AppendCloudyEdge(builder, x + width, y + height, x, y + height, 0D, -amplitude, step);
        AppendCloudyEdge(builder, x, y + height, x, y, amplitude, 0D, step);
        builder.Append("h ");
    }

    private static void AppendCloudyCirclePath(StringBuilder builder, double x, double y, double width, double height, double borderWidth, double intensity) {
        double rx = width / 2D;
        double ry = height / 2D;
        double cx = x + rx;
        double cy = y + ry;
        double amplitude = ResolveCloudyBorderAmplitude(borderWidth, intensity);
        double circumference = Math.PI * (3D * (rx + ry) - Math.Sqrt(Math.Max(0D, (3D * rx + ry) * (rx + 3D * ry))));
        int segments = Math.Max(12, (int)Math.Ceiling(circumference / ResolveCloudyBorderStep(borderWidth, intensity)));
        double startX = cx + rx;
        double startY = cy;
        builder.Append(FormatCoordinate(startX)).Append(' ').Append(FormatCoordinate(startY)).Append(" m ");
        for (int i = 0; i < segments; i++) {
            double startAngle = i * Math.PI * 2D / segments;
            double endAngle = (i + 1) * Math.PI * 2D / segments;
            double midAngle = (startAngle + endAngle) / 2D;
            double endX = cx + Math.Cos(endAngle) * rx;
            double endY = cy + Math.Sin(endAngle) * ry;
            double controlX = cx + Math.Cos(midAngle) * Math.Max(0D, rx - amplitude);
            double controlY = cy + Math.Sin(midAngle) * Math.Max(0D, ry - amplitude);
            AppendQuadraticCurve(builder, startX, startY, controlX, controlY, endX, endY);
            startX = endX;
            startY = endY;
        }

        builder.Append("h ");
    }

    private static void AppendCloudyEdge(StringBuilder builder, double startX, double startY, double endX, double endY, double normalX, double normalY, double step) {
        double dx = endX - startX;
        double dy = endY - startY;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.0001D) {
            return;
        }

        int segments = Math.Max(1, (int)Math.Ceiling(length / step));
        double currentX = startX;
        double currentY = startY;
        for (int i = 1; i <= segments; i++) {
            double t = i / (double)segments;
            double nextX = startX + dx * t;
            double nextY = startY + dy * t;
            double controlX = (currentX + nextX) / 2D + normalX;
            double controlY = (currentY + nextY) / 2D + normalY;
            AppendQuadraticCurve(builder, currentX, currentY, controlX, controlY, nextX, nextY);
            currentX = nextX;
            currentY = nextY;
        }
    }

    private static void AppendQuadraticCurve(StringBuilder builder, double startX, double startY, double controlX, double controlY, double endX, double endY) {
        double firstControlX = startX + (controlX - startX) * 2D / 3D;
        double firstControlY = startY + (controlY - startY) * 2D / 3D;
        double secondControlX = endX + (controlX - endX) * 2D / 3D;
        double secondControlY = endY + (controlY - endY) * 2D / 3D;
        builder.Append(FormatCoordinate(firstControlX)).Append(' ').Append(FormatCoordinate(firstControlY)).Append(' ')
            .Append(FormatCoordinate(secondControlX)).Append(' ').Append(FormatCoordinate(secondControlY)).Append(' ')
            .Append(FormatCoordinate(endX)).Append(' ').Append(FormatCoordinate(endY)).Append(" c ");
    }

    private static void AppendLineEnding(StringBuilder builder, string? ending, double tipX, double tipY, double oppositeX, double oppositeY, double borderWidth, PdfColor strokeColor, PdfColor fillColor) {
        if (string.IsNullOrWhiteSpace(ending) ||
            string.Equals(ending, "None", StringComparison.Ordinal)) {
            return;
        }

        double dx = tipX - oppositeX;
        double dy = tipY - oppositeY;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.0001D) {
            return;
        }

        double size = Math.Max(6D, borderWidth * 4D);
        double wing = size * 0.45D;
        double ux = dx / length;
        double uy = dy / length;
        double px = -uy;
        double py = ux;

        if (string.Equals(ending, "OpenArrow", StringComparison.Ordinal) ||
            string.Equals(ending, "ClosedArrow", StringComparison.Ordinal) ||
            string.Equals(ending, "ROpenArrow", StringComparison.Ordinal) ||
            string.Equals(ending, "RClosedArrow", StringComparison.Ordinal)) {
            bool reversed = string.Equals(ending, "ROpenArrow", StringComparison.Ordinal) ||
                string.Equals(ending, "RClosedArrow", StringComparison.Ordinal);
            double arrowTipX = reversed ? tipX - ux * size : tipX;
            double arrowTipY = reversed ? tipY - uy * size : tipY;
            double baseX = reversed ? tipX : tipX - ux * size;
            double baseY = reversed ? tipY : tipY - uy * size;
            double leftX = baseX + px * wing;
            double leftY = baseY + py * wing;
            double rightX = baseX - px * wing;
            double rightY = baseY - py * wing;

            if (string.Equals(ending, "ClosedArrow", StringComparison.Ordinal) ||
                string.Equals(ending, "RClosedArrow", StringComparison.Ordinal)) {
                builder.Append(FormatColor(fillColor)).Append(" rg ")
                    .Append(FormatCoordinate(arrowTipX)).Append(' ').Append(FormatCoordinate(arrowTipY)).Append(" m ")
                    .Append(FormatCoordinate(leftX)).Append(' ').Append(FormatCoordinate(leftY)).Append(" l ")
                    .Append(FormatCoordinate(rightX)).Append(' ').Append(FormatCoordinate(rightY)).Append(" l h B\n");
                return;
            }

            builder.Append(FormatCoordinate(leftX)).Append(' ').Append(FormatCoordinate(leftY)).Append(" m ")
                .Append(FormatCoordinate(arrowTipX)).Append(' ').Append(FormatCoordinate(arrowTipY)).Append(" l ")
                .Append(FormatCoordinate(rightX)).Append(' ').Append(FormatCoordinate(rightY)).Append(" l S\n");
            return;
        }

        if (string.Equals(ending, "Square", StringComparison.Ordinal)) {
            double half = size / 2D;
            double centerX = tipX - ux * half;
            double centerY = tipY - uy * half;
            AppendQuadrilateral(builder, fillColor, centerX, centerY, ux, uy, px, py, half, half);
            return;
        }

        if (string.Equals(ending, "Diamond", StringComparison.Ordinal)) {
            double half = size / 2D;
            double centerX = tipX - ux * half;
            double centerY = tipY - uy * half;
            builder.Append(FormatColor(fillColor)).Append(" rg ")
                .Append(FormatCoordinate(tipX)).Append(' ').Append(FormatCoordinate(tipY)).Append(" m ")
                .Append(FormatCoordinate(centerX + px * half)).Append(' ').Append(FormatCoordinate(centerY + py * half)).Append(" l ")
                .Append(FormatCoordinate(tipX - ux * size)).Append(' ').Append(FormatCoordinate(tipY - uy * size)).Append(" l ")
                .Append(FormatCoordinate(centerX - px * half)).Append(' ').Append(FormatCoordinate(centerY - py * half)).Append(" l h B\n");
            return;
        }

        if (string.Equals(ending, "Circle", StringComparison.Ordinal)) {
            double radius = size / 2D;
            double centerX = tipX - ux * radius;
            double centerY = tipY - uy * radius;
            builder.Append(FormatColor(fillColor)).Append(" rg ");
            AppendCirclePath(builder, centerX - radius, centerY - radius, size, size);
            builder.Append("B\n");
            return;
        }

        if (string.Equals(ending, "Butt", StringComparison.Ordinal)) {
            builder.Append(FormatCoordinate(tipX + px * wing)).Append(' ').Append(FormatCoordinate(tipY + py * wing)).Append(" m ")
                .Append(FormatCoordinate(tipX - px * wing)).Append(' ').Append(FormatCoordinate(tipY - py * wing)).Append(" l S\n");
            return;
        }

        if (string.Equals(ending, "Slash", StringComparison.Ordinal)) {
            double slashX = ux * wing + px * wing;
            double slashY = uy * wing + py * wing;
            builder.Append(FormatCoordinate(tipX - slashX)).Append(' ').Append(FormatCoordinate(tipY - slashY)).Append(" m ")
                .Append(FormatCoordinate(tipX + slashX)).Append(' ').Append(FormatCoordinate(tipY + slashY)).Append(" l S\n");
        }
    }

    private static void AppendQuadrilateral(StringBuilder builder, PdfColor fillColor, double centerX, double centerY, double ux, double uy, double px, double py, double halfLength, double halfWidth) {
        double frontX = centerX + ux * halfLength;
        double frontY = centerY + uy * halfLength;
        double backX = centerX - ux * halfLength;
        double backY = centerY - uy * halfLength;

        builder.Append(FormatColor(fillColor)).Append(" rg ")
            .Append(FormatCoordinate(frontX + px * halfWidth)).Append(' ').Append(FormatCoordinate(frontY + py * halfWidth)).Append(" m ")
            .Append(FormatCoordinate(backX + px * halfWidth)).Append(' ').Append(FormatCoordinate(backY + py * halfWidth)).Append(" l ")
            .Append(FormatCoordinate(backX - px * halfWidth)).Append(' ').Append(FormatCoordinate(backY - py * halfWidth)).Append(" l ")
            .Append(FormatCoordinate(frontX - px * halfWidth)).Append(' ').Append(FormatCoordinate(frontY - py * halfWidth)).Append(" l h B\n");
    }

    private static void ValidateShapeSubtype(string subtype) {
        if (!string.Equals(subtype, "Square", StringComparison.Ordinal) &&
            !string.Equals(subtype, "Circle", StringComparison.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(subtype), subtype, "PDF shape annotation subtype must be Square or Circle.");
        }
    }

    private static string FormatColor(PdfColor color) =>
        FormatCoordinate(color.R) + " " + FormatCoordinate(color.G) + " " + FormatCoordinate(color.B);

    private static string BuildBorderStrokeOperators(PdfColor color, double borderWidth, IReadOnlyList<double>? dashPattern) {
        string operators = FormatColor(color) + " RG " + FormatCoordinate(borderWidth) + " w ";
        if (dashPattern == null || dashPattern.Count == 0) {
            return operators;
        }

        var builder = new StringBuilder(operators);
        builder.Append('[');
        for (int i = 0; i < dashPattern.Count; i++) {
            if (i > 0) {
                builder.Append(' ');
            }

            builder.Append(FormatCoordinate(dashPattern[i]));
        }

        builder.Append("] 0 d ");
        return builder.ToString();
    }
}
