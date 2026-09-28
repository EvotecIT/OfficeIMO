using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Bounded ODF point-style projection and native chart style emission.</summary>
internal static class OdsChartPointStyles {
    private const int MaximumPoints = 4096;

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(XElement series,
        Func<string?, XElement?> findStyle, XElement? defaultStyle,
        IReadOnlyDictionary<string, XElement> hatches) {
        XName dataPointName = OdfNamespaces.Chart + "data-point";
        XElement[] points = series.Elements(dataPointName).Take(MaximumPoints + 1).ToArray();
        if (points.Length > MaximumPoints) return null;
        var result = new List<OfficeChartPointStyle?>();
        foreach (XElement point in points) {
            string? repeatedText = (string?)point.Attribute(OdfNamespaces.Chart + "repeated");
            int count = 1;
            if (repeatedText != null && (!int.TryParse(repeatedText, NumberStyles.None,
                    CultureInfo.InvariantCulture, out count) || count < 1)) return null;
            if (count > MaximumPoints - result.Count) return null;
            string? name = (string?)point.Attribute(OdfNamespaces.Chart + "style-name");
            OfficeChartPointStyle? appearance = null;
            if (name != null && !TryReadStyle(name, series, findStyle, defaultStyle, hatches,
                    out appearance)) return null;
            for (int index = 0; index < count; index++) result.Add(appearance);
        }
        return result.Any(style => style != null) ? result.AsReadOnly() : null;
    }

    private static bool TryReadStyle(string name, XElement series,
        Func<string?, XElement?> findStyle, XElement? defaultStyle,
        IReadOnlyDictionary<string, XElement> hatches, out OfficeChartPointStyle? appearance) {
        appearance = null;
        var attributes = new Dictionary<XName, string>();
        void AddAttributes(XElement definition) {
            XElement? graphic = definition.Element(OdfNamespaces.Style + "graphic-properties");
            if (graphic != null)
                foreach (XAttribute attribute in graphic.Attributes())
                    if (!attributes.ContainsKey(attribute.Name)) attributes.Add(attribute.Name, attribute.Value);
        }
        bool AddChain(string? styleName) {
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrEmpty(styleName)) {
                if (!visited.Add(styleName!) || visited.Count > 32) return false;
                XElement? definition = findStyle(styleName);
                if (definition == null || (string?)definition.Attribute(OdfNamespaces.Style + "family") != "chart")
                    return false;
                AddAttributes(definition);
                styleName = (string?)definition.Attribute(OdfNamespaces.Style + "parent-style-name");
            }
            return true;
        }
        if (!AddChain(name) ||
            !AddChain((string?)series.Attribute(OdfNamespaces.Chart + "style-name")) ||
            !AddChain((string?)series.Parent?.Attribute(OdfNamespaces.Chart + "style-name"))) return false;
        if (defaultStyle != null) AddAttributes(defaultStyle);
        string? Get(XName key) => attributes.TryGetValue(key, out string? value) ? value : null;
        if (Get(OdfNamespaces.Draw + "fill-opacity") != null ||
            Get(OdfNamespaces.Draw + "fill-transparency") != null ||
            Get(OdfNamespaces.Svg + "stroke-opacity") != null ||
            Get(OdfNamespaces.Draw + "stroke-dash") != null) return false;
        string? fillMode = Get(OdfNamespaces.Draw + "fill");
        OfficeColor? fill = null;
        bool noFill = fillMode == "none";
        OfficeChartHatchPattern? hatch = null;
        OfficeColor? hatchColor = null;
        if (fillMode != null && fillMode is not ("none" or "solid" or "hatch")) return false;
        string? fillText = Get(OdfNamespaces.Draw + "fill-color");
        if (fillText != null && !OfficeColor.TryParseHex(fillText, out _)) return false;
        if (fillMode == "hatch") {
            if (Get(OdfNamespaces.Draw + "fill-hatch-solid") != "true" ||
                !OfficeColor.TryParseHex(fillText, out OfficeColor background) ||
                !TryReadHatch(Get(OdfNamespaces.Draw + "fill-hatch-name"), hatches,
                    out OfficeChartHatchPattern pattern, out OfficeColor ink)) return false;
            fill = background;
            hatch = pattern;
            hatchColor = ink;
        } else if (!noFill && fillText != null) {
            OfficeColor.TryParseHex(fillText, out OfficeColor color);
            fill = color;
        } else if (fillMode == "solid") return false;
        string? strokeMode = Get(OdfNamespaces.Draw + "stroke");
        if (strokeMode != null && strokeMode is not ("none" or "solid")) return false;
        bool? showOutline = strokeMode == "none" ? false : strokeMode == "solid" ? true : null;
        OfficeColor? outline = null;
        string? outlineText = Get(OdfNamespaces.Svg + "stroke-color");
        if (outlineText != null) {
            if (!OfficeColor.TryParseHex(outlineText, out OfficeColor parsed)) return false;
            outline = parsed;
        }
        double? width = null;
        string? widthText = Get(OdfNamespaces.Svg + "stroke-width");
        if (widthText != null) {
            if (string.IsNullOrWhiteSpace(widthText) ||
                !OdfLength.Parse(widthText).TryToPoints(out double points) || points <= 0 || points > 1584)
                return false;
            width = points;
        }
        OfficeStrokeLineJoin? join = Get(OdfNamespaces.Draw + "stroke-linejoin") switch {
            "miter" => OfficeStrokeLineJoin.Miter,
            "round" => OfficeStrokeLineJoin.Round,
            "bevel" => OfficeStrokeLineJoin.Bevel,
            null => null,
            _ => (OfficeStrokeLineJoin)(-1)
        };
        if (join.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineJoin), join.Value)) return false;
        if (fill == null && !noFill && hatch == null && showOutline == null && outline == null && width == null && join == null)
            return true;
        appearance = new OfficeChartPointStyle(fill, noFill, hatch, hatchColor, outline, width, showOutline, join);
        return true;
    }

    internal static IReadOnlyDictionary<string, XElement> IndexHatches(XDocument content, XDocument? styles) {
        var index = new Dictionary<string, XElement>(StringComparer.Ordinal);
        void Add(XDocument part) {
            foreach (XElement definition in part.Descendants(OdfNamespaces.Draw + "hatch")) {
                string? name = (string?)definition.Attribute(OdfNamespaces.Draw + "name");
                if (name != null && !index.ContainsKey(name)) index.Add(name, definition);
            }
        }
        Add(content);
        if (styles != null) Add(styles);
        return index;
    }

    private static bool TryReadHatch(string? name, IReadOnlyDictionary<string, XElement> hatches,
        out OfficeChartHatchPattern pattern, out OfficeColor color) {
        pattern = default;
        color = default;
        if (string.IsNullOrEmpty(name) || !hatches.TryGetValue(name!, out XElement? definition) ||
            !OfficeColor.TryParseHex((string?)definition.Attribute(OdfNamespaces.Draw + "color"), out color) ||
            !int.TryParse((string?)definition.Attribute(OdfNamespaces.Draw + "rotation"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int rotation)) return false;
        string? distanceText = (string?)definition.Attribute(OdfNamespaces.Draw + "distance");
        if (string.IsNullOrWhiteSpace(distanceText) ||
            !OdfLength.Parse(distanceText!).TryToPoints(out double spacing)) return false;
        string? kind = (string?)definition.Attribute(OdfNamespaces.Draw + "style");
        int angle = ((rotation % 1800) + 1800) % 1800;
        pattern = kind switch {
            "single" when angle == 0 => OfficeChartHatchPattern.Vertical,
            "single" when angle == 900 => OfficeChartHatchPattern.Horizontal,
            "single" when angle == 450 => OfficeChartHatchPattern.ForwardDiagonal,
            "single" when angle == 1350 => OfficeChartHatchPattern.BackwardDiagonal,
            "double" when angle == 0 || angle == 900 => OfficeChartHatchPattern.Cross,
            "double" when angle == 450 || angle == 1350 => OfficeChartHatchPattern.DiagonalCross,
            _ => (OfficeChartHatchPattern)(-1)
        };
        if (!Enum.IsDefined(typeof(OfficeChartHatchPattern), pattern)) return false;
        double normalSpacing = OdfLength.Centimeters(0.15).ToPoints();
        double wideSpacing = OdfLength.Centimeters(0.3).ToPoints();
        if (Math.Abs(spacing - normalSpacing) <= 0.05D) return true;
        if (pattern == OfficeChartHatchPattern.ForwardDiagonal &&
            Math.Abs(spacing - wideSpacing) <= 0.05D) {
            pattern = OfficeChartHatchPattern.WideForwardDiagonal;
            return true;
        }
        return false;
    }

    internal static void Write(IReadOnlyList<OfficeChartPointStyle?>? pointStyles, int pointCount,
        int seriesIndex, XElement series, ICollection<XElement> styleDefinitions,
        ICollection<XElement> hatchDefinitions) {
        if (pointStyles == null || pointStyles.All(style => style == null)) {
            series.Add(new XElement(OdfNamespaces.Chart + "data-point",
                new XAttribute(OdfNamespaces.Chart + "repeated", pointCount)));
            return;
        }
        if (pointStyles.Count != pointCount)
            throw new ArgumentException("Point-style count must match the chart category count.", nameof(pointStyles));
        for (int index = 0; index < pointCount; index++) {
            OfficeChartPointStyle? point = pointStyles[index];
            var dataPoint = new XElement(OdfNamespaces.Chart + "data-point");
            if (point != null) {
                string styleName = "ChartPoint" + seriesIndex.ToString(CultureInfo.InvariantCulture) + "_" + index.ToString(CultureInfo.InvariantCulture);
                var graphic = new XElement(OdfNamespaces.Style + "graphic-properties");
                if (point.NoFill) graphic.SetAttributeValue(OdfNamespaces.Draw + "fill", "none");
                else if (point.Hatch.HasValue) {
                    string hatchName = "ChartHatch" + seriesIndex.ToString(CultureInfo.InvariantCulture) + "_" + index.ToString(CultureInfo.InvariantCulture);
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill", "hatch");
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill-hatch-name", hatchName);
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill-hatch-solid", "true");
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill-color", point.FillColor!.Value.ToString());
                    hatchDefinitions.Add(CreateHatch(hatchName, point));
                } else if (point.FillColor.HasValue) {
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill", "solid");
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "fill-color", point.FillColor.Value.ToString());
                }
                if (point.ShowOutline == false) graphic.SetAttributeValue(OdfNamespaces.Draw + "stroke", "none");
                else if (point.ShowOutline == true || point.OutlineColor.HasValue || point.OutlineWidth.HasValue || point.OutlineJoin.HasValue)
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "stroke", "solid");
                if (point.OutlineColor.HasValue)
                    graphic.SetAttributeValue(OdfNamespaces.Svg + "stroke-color", point.OutlineColor.Value.ToString());
                if (point.OutlineWidth.HasValue)
                    graphic.SetAttributeValue(OdfNamespaces.Svg + "stroke-width", OdfLength.Points(point.OutlineWidth.Value).ToString());
                if (point.OutlineJoin.HasValue)
                    graphic.SetAttributeValue(OdfNamespaces.Draw + "stroke-linejoin",
                        point.OutlineJoin.Value.ToString().ToLowerInvariant());
                styleDefinitions.Add(new XElement(OdfNamespaces.Style + "style",
                    new XAttribute(OdfNamespaces.Style + "name", styleName),
                    new XAttribute(OdfNamespaces.Style + "family", "chart"), graphic));
                dataPoint.SetAttributeValue(OdfNamespaces.Chart + "style-name", styleName);
            }
            series.Add(dataPoint);
        }
    }

    internal static void Validate(IReadOnlyList<OfficeChartPointStyle?>? styles) {
        if (styles == null) return;
        foreach (OfficeChartPointStyle? style in styles) {
            if (style == null) continue;
            foreach (OfficeColor? color in new[] { style.FillColor, style.HatchColor, style.OutlineColor })
                if (color.HasValue && color.Value.A != 255)
                    throw new NotSupportedException("ODF chart point colors require opaque RGB values.");
        }
    }

    private static XElement CreateHatch(string name, OfficeChartPointStyle point) {
        OfficeChartHatchPattern pattern = point.Hatch!.Value;
        bool crossed = pattern is OfficeChartHatchPattern.Cross or OfficeChartHatchPattern.DiagonalCross;
        int rotation = pattern switch {
            OfficeChartHatchPattern.Vertical or OfficeChartHatchPattern.Cross => 0,
            OfficeChartHatchPattern.Horizontal => 900,
            OfficeChartHatchPattern.ForwardDiagonal or OfficeChartHatchPattern.WideForwardDiagonal or OfficeChartHatchPattern.DiagonalCross => 450,
            _ => 1350
        };
        return new XElement(OdfNamespaces.Draw + "hatch",
            new XAttribute(OdfNamespaces.Draw + "name", name),
            new XAttribute(OdfNamespaces.Draw + "style", crossed ? "double" : "single"),
            new XAttribute(OdfNamespaces.Draw + "color", point.HatchColor!.Value.ToString()),
            new XAttribute(OdfNamespaces.Draw + "distance", pattern == OfficeChartHatchPattern.WideForwardDiagonal ? "0.3cm" : "0.15cm"),
            new XAttribute(OdfNamespaces.Draw + "rotation", rotation));
    }
}
