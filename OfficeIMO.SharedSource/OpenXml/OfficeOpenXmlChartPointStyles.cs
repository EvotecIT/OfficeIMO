using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native chart point-style codec; Core remains free of OpenXML dependencies.</summary>
internal static class OfficeOpenXmlChartPointStyles {
    internal static bool IsSupportedPattern(OfficeXmlNode pattern, OfficeXmlNode scheme) =>
        ReadHatch(pattern.Attribute("prst")).HasValue &&
        !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.Child(NsA, "fgClr")) &&
        !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.Child(NsA, "bgClr")) &&
        OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "fgClr"), scheme).HasValue &&
        OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "bgClr"), scheme).HasValue;

    // Preserve each format's point-override budget independently of its cache limit.
    // Count duplicate and invalid records before interpreting appearances.
    internal static IReadOnlyList<OfficeXmlNode> GetBoundedPoints(OfficeXmlNode series, int maximum = 1_000_000) {
        var points = new List<OfficeXmlNode>();
        foreach (OfficeXmlNode point in series.Children(NsC, "dPt")) {
            points.Add(point);
            if (points.Count > maximum)
                throw new System.IO.InvalidDataException($"The chart exceeds the supported limit of {maximum} point overrides.");
        }
        return points;
    }

    internal static bool IsSupported(OfficeXmlNode properties, OfficeXmlNode scheme, bool flattenThreeDimensional = false) {
        if (properties.HasAttributes) return false;
        foreach (OfficeXmlNode child in properties.Children()) {
            string name = child.NamespaceId == NsA ? child.LocalName : string.Empty;
            if (name == "effectLst" && !child.HasChildElements() && !child.HasAttributes) continue;
            if (name == "sp3d" && !child.HasChildElements() && !child.HasAttributes) continue;
            if (flattenThreeDimensional && name == "sp3d") continue;
            if (name == "noFill") continue;
            if (name == "solidFill") {
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(child) ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme).HasValue) return false;
                continue;
            }
            if (name == "pattFill") {
                if (!IsSupportedPattern(child, scheme)) return false;
                continue;
            }
            if (name == "ln") {
                int? width = child.IntAttr("w");
                if (width == 0 && child.Child(NsA, "noFill").IsNone) return false;
                if (width < 0 || width > 20116800 ||
                    HasUnsupportedLineAttributes(child)) return false;
                foreach (OfficeXmlNode lineChild in child.Children()) {
                    string lineName = lineChild.NamespaceId == NsA ? lineChild.LocalName : string.Empty;
                    if (lineName == "noFill") continue;
                    if (lineName == "solidFill" && !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(lineChild) &&
                        OfficeOpenXmlThemeColorResolver.ResolveColor(lineChild, scheme).HasValue) continue;
                    if (lineName is "round" or "bevel") continue;
                    if (lineName == "miter" && !lineChild.HasAttributes && !lineChild.HasChildElements()) continue;
                    return false;
                }
                continue;
            }
            return false;
        }
        return true;
    }

    // Line caps are drawn on open series lines and have no visible effect on closed shapes,
    // so only pen alignment and compound lines change what a chart outline looks like.
    internal static bool HasUnsupportedLineAttributes(OfficeXmlNode outline) =>
        outline.Attribute("algn") is string alignment && alignment != "ctr" ||
        outline.Attribute("cmpd") is string compound && compound != "sng";

    internal static bool HasUnsupportedPointContent(OfficeXmlNode point) {
        OfficeXmlNode direct = point.Child(NsC, "spPr");
        bool sawMarker = false;
        foreach (OfficeXmlNode child in point.Children()) {
            string name = child.NamespaceId == NsC ? child.LocalName : string.Empty;
            if (name is "idx" or "spPr") continue;
            if (name == "explosion" && IsPieSeries(point.Parent)) continue;
            if (name == "marker" && !sawMarker && !child.HasAttributes &&
                child.Children().ToList().TrueForAll(markerChild => markerChild.Is(NsC, "spPr")) &&
                (child.Child(NsC, "spPr").IsNone || direct.IsNone ||
                    !direct.HasChildElements() && !direct.HasAttributes)) {
                sawMarker = true;
                continue;
            }
            if (name == "invertIfNegative" && child.BoolVal() == false) continue;
            if (name == "bubble3D" && child.BoolVal() == false) continue;
            if (name == "extLst" && HasOnlyUniqueIdMetadata(child)) continue;
            return true;
        }
        return false;
    }

    private static bool HasOnlyUniqueIdMetadata(OfficeXmlNode extensions) {
        OfficeXmlNode extension = extensions.FirstChild;
        if (extensions.Children().Count != 1 || !extension.Is(NsC, "ext") ||
            extension.Attribute("uri") != "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" ||
            extension.Children().Count != 1) return false;
        OfficeXmlNode uniqueId = extension.FirstChild;
        if (uniqueId.LocalName != "uniqueId" ||
            uniqueId.NamespaceUri != C16Uri ||
            uniqueId.HasChildElements()) return false;
        List<OfficeXmlAttribute> attributes = uniqueId.Attributes();
        return attributes.Count == 1 && attributes[0].LocalName == "val" &&
            attributes[0].NamespaceUri.Length == 0 && Guid.TryParse(attributes[0].Value, out _);
    }

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(OfficeXmlNode series, int count, OfficeXmlNode scheme) =>
        Read(GetBoundedPoints(series), count, scheme, series.Parent.LocalName.EndsWith("3DChart", StringComparison.Ordinal));

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(IReadOnlyList<OfficeXmlNode> points, int count, OfficeXmlNode scheme, bool flattenThreeDimensional = false) {
        OfficeChartPointStyle?[]? styles = null;
        foreach (OfficeXmlNode point in points) {
            uint? index = point.Child(NsC, "idx").UIntVal();
            if (!index.HasValue || index.Value >= (uint)count) continue;
            if (HasUnsupportedPointContent(point))
                throw new System.IO.InvalidDataException("The native point record contains unsupported appearance or behavior.");
            OfficeXmlNode properties = point.Child(NsC, "marker").Child(NsC, "spPr");
            if (properties.IsNone) properties = point.Child(NsC, "spPr");
            if (properties.IsNone) continue;
            if (!IsSupported(properties, scheme, flattenThreeDimensional))
                throw new System.IO.InvalidDataException("The native point appearance cannot be projected by the shared chart model.");
            var style = ReadStyle(properties, scheme);
            if (style == null) continue;
            styles ??= new OfficeChartPointStyle?[count];
            styles[(int)index.Value] = style;
        }
        return styles;
    }

    internal static OfficeChartPointStyle? ReadSeriesStyle(OfficeXmlNode series, OfficeXmlNode scheme) {
        var style = ReadStyle(series.Child(NsC, "spPr"), scheme);
        if (style == null || !OfficeOpenXmlChartAutomaticColors.HasAutomaticSeriesColor(series) || style.Hatch.HasValue) return style;
        // The writer stores a concrete default colour for native Office. Keep that
        // colour on the series, so a caller palette can still replace it; do not
        // turn the stored default into a second explicit paint override.
        OfficeColor? automatic = style.FillColor ?? style.OutlineColor;
        return new OfficeChartPointStyle(null, style.NoFill, outlineColor: style.OutlineColor == automatic ? null : style.OutlineColor,
            outlineWidth: style.OutlineWidth, showOutline: style.ShowOutline, outlineJoin: style.OutlineJoin).WithOutlineCap(style.OutlineCap);
    }

    /// <summary>Projects qualified ribbon paint into the shared flat stroke without confusing its fill and border.</summary>
    internal static (OfficeChartPointStyle Paint, OfficeColor? Color, bool Automatic) ReadRibbonPaint(
        OfficeXmlNode series, OfficeXmlNode scheme, bool useSpreadsheetDefaultTheme = false) =>
        ReadRibbonPaint(series, scheme, new OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext(),
            useSpreadsheetDefaultTheme);

    internal static (OfficeChartPointStyle Paint, OfficeColor? Color, bool Automatic) ReadRibbonPaint(
        OfficeXmlNode series, OfficeXmlNode scheme,
        OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext colorContext,
        bool useSpreadsheetDefaultTheme = false) {
        var native = ReadStyle(series.Child(NsC, "spPr"), scheme);
        var inherited = ReadSeriesStyle(series, scheme);
        bool fill = native?.NoFill != true, outline = native?.ShowOutline != false;
        OfficeColor? automatic = (fill && native?.FillColor == null || outline && native?.OutlineColor == null)
            ? OfficeOpenXmlChartAutomaticColors.ResolveSeries(series, scheme, colorContext,
                useSpreadsheetDefaultTheme) : null;
        OfficeColor? body = native?.FillColor ?? (fill ? automatic : null);
        OfficeColor? border = native?.OutlineColor ?? (outline ? automatic : null);
        bool visibleFill = fill && (!body.HasValue || body.Value.A > 0);
        bool visibleOutline = outline && (!border.HasValue || border.Value.A > 0);
        OfficeColor? color = visibleFill ? body : visibleOutline ? border : body ?? border;
        bool automaticColor = OfficeOpenXmlChartAutomaticColors.HasAutomaticSeriesColor(series) ||
            (visibleFill ? native?.FillColor == null : visibleOutline && native?.OutlineColor == null);
        var paint = new OfficeChartPointStyle(inherited?.FillColor, native?.NoFill ?? false,
            inherited?.Hatch, inherited?.HatchColor, automaticColor ? null : color,
            native?.OutlineWidth, visibleFill || visibleOutline, native?.OutlineJoin);
        return (paint, color, automaticColor);
    }

    internal static OfficeChartPointStyle? ReadStyle(OfficeXmlNode properties, OfficeXmlNode scheme) {
        if (properties.IsNone) return null;
        OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.Child(NsA, "solidFill"), scheme);
        bool noFill = properties.Child(NsA, "noFill").Exists;
        OfficeXmlNode pattern = properties.Child(NsA, "pattFill");
        OfficeChartHatchPattern? hatch = ReadHatch(pattern.Attribute("prst"));
        OfficeColor? hatchColor = hatch.HasValue ? OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "fgClr"), scheme) : null;
        if (hatch.HasValue && hatchColor.HasValue)
            fill = OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.Child(NsA, "bgClr"), scheme);
        else { hatch = null; hatchColor = null; }
        OfficeXmlNode outline = properties.Child(NsA, "ln");
        OfficeColor? line = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), scheme);
        bool? showOutline = outline.IsNone ? null : outline.Child(NsA, "noFill").Exists ? false : true;
        OfficeLength? width = outline.IntAttr("w") is int emu && emu > 0 ? OfficeLength.Emus(emu) : null;
        OfficeStrokeLineJoin? join = outline.Child(NsA, "round").Exists ? OfficeStrokeLineJoin.Round :
            outline.Child(NsA, "bevel").Exists ? OfficeStrokeLineJoin.Bevel :
            outline.Child(NsA, "miter").Exists ? OfficeStrokeLineJoin.Miter : null;
        string? capType = outline.Attribute("cap");
        OfficeStrokeLineCap? cap = capType == "rnd" ? OfficeStrokeLineCap.Round :
            capType == "sq" ? OfficeStrokeLineCap.Square : null;
        if (!fill.HasValue && !noFill && !hatch.HasValue && !line.HasValue && !width.HasValue && !showOutline.HasValue && !join.HasValue &&
            !cap.HasValue) return null;

        return new OfficeChartPointStyle(fill, noFill, hatch, hatchColor, line, width, showOutline, join).WithOutlineCap(cap);
    }

    internal static void ApplySeries(OfficeXmlNode series, OfficeSeries data, OfficeTheme? theme = null) {
        if (data.SeriesStyle != null) {
            OfficeXmlNode properties = series.Child(NsC, "spPr");
            bool created = properties.IsNone;
            if (created) properties = series.CreateElement(NsC, "spPr");
            ApplySeriesPaint(properties, data.SeriesStyle, theme);
            if (created) SetOrdered(series, properties, SeriesOrder(series));
        }
        if (data.PointStyles == null) return;
        ApplyPoints(series, data.PointStyles, data.PointColors, data.Values.Count, theme);
    }

    internal static void ApplyPoints(OfficeXmlNode series,
        IReadOnlyList<OfficeChartPointStyle?>? styles, IReadOnlyList<OfficeColor?>? colors, int count, OfficeTheme? theme = null) {
        if (styles == null && colors == null) return;
        var existing = new Dictionary<uint, OfficeXmlNode>();
        foreach (OfficeXmlNode point in GetBoundedPoints(series)) {
            if (point.Child(NsC, "idx").UIntVal() is uint index && !existing.ContainsKey(index))
                existing.Add(index, point);
        }
        OfficeXmlNode anchor = FindPointAnchor(series);
        for (int index = 0; index < count; index++) {
            OfficeChartPointStyle? style = styles?[index];
            OfficeColor? color = colors?[index];
            existing.TryGetValue((uint)index, out OfficeXmlNode point);
            if (point.IsNone && style == null && !color.HasValue) continue;
            ApplyPointCore(series, (uint)index, style, color, point, anchor, theme);
        }
    }

    internal static void ApplyPoint(OfficeXmlNode series, uint index, OfficeChartPointStyle? style, OfficeColor? legacyColor = null) {
        if (style?.OutlineWidth > OfficeLength.Points(1584)) throw new ArgumentOutOfRangeException(nameof(style), "Native chart outlines cannot exceed 1584 points.");
        if (!SupportsDataPoints(series))
            throw new NotSupportedException("Point styles are not supported by this native chart series.");
        OfficeXmlNode point = OfficeXmlNode.None;
        foreach (OfficeXmlNode item in GetBoundedPoints(series)) {
            if (item.Child(NsC, "idx").UIntVal() == index) {
                point = item;
                break;
            }
        }
        if (point.IsNone && style == null && !legacyColor.HasValue) return;
        ApplyPointCore(series, index, style, legacyColor, point, FindPointAnchor(series));
    }

    private static OfficeXmlNode FindPointAnchor(OfficeXmlNode series) {
        foreach (OfficeXmlNode child in series.Children()) {
            if (child.NamespaceId == NsC && child.LocalName is "dLbls" or "trendline" or "errBars" or "cat" or "val" or
                "xVal" or "yVal" or "bubbleSize" or "smooth" or "extLst") return child;
        }
        return OfficeXmlNode.None;
    }

    private static bool SupportsDataPoints(OfficeXmlNode series) =>
        series.Is(NsC, "ser") && series.Parent.NamespaceId == NsC && series.Parent.LocalName is
            "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or "stockChart" or
            "areaChart" or "area3DChart" or "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart" or
            "scatterChart" or "radarChart" or "bubbleChart";

    private static bool IsPieSeries(OfficeXmlNode series) =>
        series.Is(NsC, "ser") && series.Parent.NamespaceId == NsC &&
        series.Parent.LocalName is "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart";

    /// <summary>The schema order of a <c>c:ser</c>'s children, which depends on the chart group that holds it.</summary>
    internal static OfficeXmlChildOrder SeriesOrder(OfficeXmlNode series) => series.Parent.LocalName switch {
        "barChart" or "bar3DChart" => CBarSerView.Order,
        "lineChart" or "line3DChart" or "stockChart" => CLineSerView.Order,
        "areaChart" or "area3DChart" => CAreaSerView.Order,
        "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart" => CPieSerView.Order,
        "scatterChart" => CScatterSerView.Order,
        "radarChart" => CRadarSerView.Order,
        "bubbleChart" => CBubbleSerView.Order,
        "surfaceChart" or "surface3DChart" => CSurfaceSerView.Order,
        _ => throw new InvalidOperationException("The series has no chart group that defines its content.")
    };

    private static void ApplyPointCore(OfficeXmlNode series, uint index,
        OfficeChartPointStyle? style, OfficeColor? legacyColor, OfficeXmlNode point, OfficeXmlNode anchor, OfficeTheme? theme = null) {
        if (style?.OutlineWidth > OfficeLength.Points(1584)) throw new ArgumentOutOfRangeException(nameof(style), "Native chart outlines cannot exceed 1584 points.");
        if (point.IsNone) {
            point = series.CreateElement(NsC, "dPt");
            point.AppendVal(NsC, "idx", Text(index));
            point = series.InsertBefore(point, anchor);
        }
        ApplyPointStyle(point, style, legacyColor, theme);
    }

    private static void ApplyPointStyle(OfficeXmlNode point, OfficeChartPointStyle? style, OfficeColor? legacyColor, OfficeTheme? theme) {
        OfficeXmlNode pointMarker = point.Child(NsC, "marker");
        bool ownerIsMarker = pointMarker.Child(NsC, "spPr").Exists;
        OfficeXmlNode styleOwner = ownerIsMarker ? pointMarker : point;
        OfficeXmlNode properties = styleOwner.Child(NsC, "spPr");
        bool created = properties.IsNone;
        if (created) properties = styleOwner.CreateElement(NsC, "spPr");
        OfficeXmlNode competing = !ownerIsMarker
            ? pointMarker.Child(NsC, "spPr")
            : point.Child(NsC, "spPr");
        if (competing.Exists) {
            ClearPointAppearance(competing);
            if (!competing.HasChildElements() && !competing.HasAttributes) competing.Remove();
        }
        ApplyPaint(properties, style, legacyColor, theme);
        bool hasContent = properties.HasChildElements() || properties.HasAttributes;
        if (created && hasContent) SetOrdered(styleOwner, properties, ownerIsMarker ? CMarkerView.Order : CDPtView.Order);
        else if (!hasContent && !created) properties.Remove();
    }

    internal static void ApplySeriesPaint(OfficeXmlNode properties, OfficeChartPointStyle style, OfficeTheme? theme = null) {
        // A detached copy to build the requested paint in, emptied first; only its children are used.
        OfficeXmlNode replacement = properties.Clone();
        replacement.RemoveAllChildElements();
        ApplyPaint(replacement, style, null, theme);
        if (style.NoFill || style.FillColor.HasValue || style.Hatch.HasValue) {
            RemoveFills(properties);
            foreach (OfficeXmlNode child in replacement.Children().ToList()) {
                if (child.Is(NsA, "ln")) continue;
                child.Remove();
                SetOrdered(properties, child, DmlShapePropertiesView.Order);
            }
        }
        OfficeXmlNode requested = replacement.Child(NsA, "ln");
        if (requested.Exists) {
            OfficeXmlNode outline = properties.Child(NsA, "ln");
            bool created = outline.IsNone;
            if (created) outline = properties.CreateElement(NsA, "ln");
            if (requested.Attribute("w") is string width) outline.SetAttribute("w", width);
            if (requested.Attribute("cap") is string cap) outline.SetAttribute("cap", cap);
            if (style.ShowOutline == false || style.OutlineColor.HasValue) {
                RemoveFills(outline);
            } else if (style.ShowOutline == true) outline.RemoveChildren(NsA, "noFill");
            if (style.OutlineJoin.HasValue) {
                outline.RemoveChildren(NsA, "round");
                outline.RemoveChildren(NsA, "bevel");
                outline.RemoveChildren(NsA, "miter");
            }
            foreach (OfficeXmlNode child in requested.Children().ToList()) {
                child.Remove();
                SetOrdered(outline, child, DmlLinePropertiesView.Order);
            }
            if (created) SetOrdered(properties, outline, DmlShapePropertiesView.Order);
        }
    }

    // The fill choices of a:spPr and a:ln.
    private static void RemoveFills(OfficeXmlNode element) {
        element.RemoveChildren(NsA, "solidFill");
        element.RemoveChildren(NsA, "noFill");
        element.RemoveChildren(NsA, "pattFill");
        element.RemoveChildren(NsA, "gradFill");
        element.RemoveChildren(NsA, "blipFill");
        element.RemoveChildren(NsA, "grpFill");
    }

    private static void ApplyPaint(OfficeXmlNode properties, OfficeChartPointStyle? style, OfficeColor? legacyColor, OfficeTheme? theme) {
        ClearPointAppearance(properties);
        if (style?.NoFill == true) SetOrdered(properties, properties.CreateElement(NsA, "noFill"), DmlShapePropertiesView.Order);
        else if (style?.Hatch is OfficeChartHatchPattern hatch) {
            OfficeXmlNode pattern = properties.CreateElement(NsA, "pattFill");
            pattern.SetAttribute("prst", HatchToken(hatch));
            AppendColor(pattern.AppendNew(NsA, "fgClr"), style.HatchColor!.Value, theme);
            AppendColor(pattern.AppendNew(NsA, "bgClr"), style.FillColor ?? legacyColor ?? OfficeColor.White, theme);
            SetOrdered(properties, pattern, DmlShapePropertiesView.Order);
        } else if ((style?.FillColor ?? legacyColor) is OfficeColor fill) {
            OfficeXmlNode solid = properties.CreateElement(NsA, "solidFill");
            AppendColor(solid, fill, theme);
            SetOrdered(properties, solid, DmlShapePropertiesView.Order);
        }
        if (style?.ShowOutline == false) {
            OfficeXmlNode hidden = properties.CreateElement(NsA, "ln");
            hidden.AppendNew(NsA, "noFill");
            SetOrdered(properties, hidden, DmlShapePropertiesView.Order);
        } else if (style?.OutlineColor != null || style?.OutlineWidth != null || style?.ShowOutline == true || style?.OutlineJoin != null ||
            style?.OutlineCap != null) {
            OfficeXmlNode outline = properties.CreateElement(NsA, "ln");
            if (style.OutlineWidth.HasValue) outline.SetAttribute("w", Text(checked((int)style.OutlineWidth.Value.ToEmus())));
            if (style.OutlineCap is OfficeStrokeLineCap cap)
                outline.SetAttribute("cap", cap == OfficeStrokeLineCap.Round ? "rnd" :
                    cap == OfficeStrokeLineCap.Square ? "sq" : "flat");
            if (style.OutlineColor is OfficeColor line) AppendColor(outline.AppendNew(NsA, "solidFill"), line, theme);
            if (style.OutlineJoin is OfficeStrokeLineJoin join) outline.AppendNew(NsA, join switch {
                OfficeStrokeLineJoin.Round => "round",
                OfficeStrokeLineJoin.Bevel => "bevel",
                _ => "miter"
            });
            SetOrdered(properties, outline, DmlShapePropertiesView.Order);
        }
    }

    private static void ClearPointAppearance(OfficeXmlNode properties) {
        RemoveFills(properties);
        properties.RemoveChildren(NsA, "ln");
    }

    // Appends the colour as a:srgbClr, with its alpha when it is not opaque.
    private static void AppendColor(OfficeXmlNode parent, OfficeColor color, OfficeTheme? theme) {
        color = color.Resolve(theme ?? OfficeTheme.Default);
        OfficeXmlNode value = parent.AppendVal(NsA, "srgbClr", color.ToRgbHex());
        if (color.A != 255) value.AppendVal(NsA, "alpha", Text((int)Math.Round(color.A * 100000D / 255)));
    }

    internal static OfficeChartHatchPattern? ReadHatch(string? token) => token switch {
        "horz" => OfficeChartHatchPattern.Horizontal, "vert" => OfficeChartHatchPattern.Vertical,
        "upDiag" => OfficeChartHatchPattern.ForwardDiagonal, "dnDiag" => OfficeChartHatchPattern.BackwardDiagonal,
        "cross" => OfficeChartHatchPattern.Cross, "diagCross" => OfficeChartHatchPattern.DiagonalCross,
        "wdUpDiag" => OfficeChartHatchPattern.WideForwardDiagonal, _ => null
    };
    private static string HatchToken(OfficeChartHatchPattern hatch) => hatch switch {
        OfficeChartHatchPattern.Horizontal => "horz", OfficeChartHatchPattern.Vertical => "vert",
        OfficeChartHatchPattern.ForwardDiagonal => "upDiag", OfficeChartHatchPattern.BackwardDiagonal => "dnDiag",
        OfficeChartHatchPattern.Cross => "cross", OfficeChartHatchPattern.DiagonalCross => "diagCross",
        OfficeChartHatchPattern.WideForwardDiagonal => "wdUpDiag",
        _ => throw new ArgumentOutOfRangeException(nameof(hatch))
    };

    // SDK callers
    internal static bool IsSupportedPattern(A.PatternFill pattern, A.ColorScheme? scheme) => IsSupportedPattern(OfficeXmlNode.FromSdk(pattern), OfficeXmlNode.FromSdk(scheme));
    internal static IReadOnlyList<C.DataPoint> GetBoundedPoints(OpenXmlElement series, int maximum = 1_000_000) =>
        GetBoundedPoints(OfficeXmlNode.FromSdk(series), maximum).Select(point => (C.DataPoint)point.SdkElement!).ToList();
    internal static bool IsSupported(C.ChartShapeProperties properties, A.ColorScheme? scheme, bool flattenThreeDimensional = false) =>
        IsSupported(OfficeXmlNode.FromSdk(properties), OfficeXmlNode.FromSdk(scheme), flattenThreeDimensional);
    internal static bool HasUnsupportedLineAttributes(A.Outline outline) => HasUnsupportedLineAttributes(OfficeXmlNode.FromSdk(outline));
    internal static bool HasUnsupportedPointContent(C.DataPoint point) => HasUnsupportedPointContent(OfficeXmlNode.FromSdk(point));
    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(OpenXmlElement series, int count, A.ColorScheme? scheme) =>
        Read(OfficeXmlNode.FromSdk(series), count, OfficeXmlNode.FromSdk(scheme));
    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(IReadOnlyList<C.DataPoint> points, int count, A.ColorScheme? scheme, bool flattenThreeDimensional = false) =>
        Read(points.Select(point => OfficeXmlNode.FromSdk(point)).ToList(), count, OfficeXmlNode.FromSdk(scheme), flattenThreeDimensional);
    internal static OfficeChartPointStyle? ReadSeriesStyle(OpenXmlElement series, A.ColorScheme? scheme) => ReadSeriesStyle(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(scheme));
    internal static (OfficeChartPointStyle Paint, OfficeColor? Color, bool Automatic) ReadRibbonPaint(OpenXmlElement series, A.ColorScheme? scheme, bool useSpreadsheetDefaultTheme = false) =>
        ReadRibbonPaint(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(scheme), useSpreadsheetDefaultTheme);
    internal static (OfficeChartPointStyle Paint, OfficeColor? Color, bool Automatic) ReadRibbonPaint(OpenXmlElement series, A.ColorScheme? scheme, OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext colorContext, bool useSpreadsheetDefaultTheme = false) =>
        ReadRibbonPaint(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(scheme), colorContext, useSpreadsheetDefaultTheme);
    internal static OfficeChartPointStyle? ReadStyle(C.ChartShapeProperties? properties, A.ColorScheme? scheme) => ReadStyle(OfficeXmlNode.FromSdk(properties), OfficeXmlNode.FromSdk(scheme));
    internal static void ApplySeries(OpenXmlCompositeElement series, OfficeSeries data, OfficeTheme? theme = null) => ApplySeries(OfficeXmlNode.FromSdk(series), data, theme);
    internal static void ApplyPoints(OpenXmlCompositeElement series, IReadOnlyList<OfficeChartPointStyle?>? styles, IReadOnlyList<OfficeColor?>? colors, int count, OfficeTheme? theme = null) =>
        ApplyPoints(OfficeXmlNode.FromSdk(series), styles, colors, count, theme);
    internal static void ApplyPoint(OpenXmlCompositeElement series, uint index, OfficeChartPointStyle? style, OfficeColor? legacyColor = null) =>
        ApplyPoint(OfficeXmlNode.FromSdk(series), index, style, legacyColor);
    internal static void ApplySeriesPaint(C.ChartShapeProperties properties, OfficeChartPointStyle style, OfficeTheme? theme = null) =>
        ApplySeriesPaint(OfficeXmlNode.FromSdk(properties), style, theme);
}
