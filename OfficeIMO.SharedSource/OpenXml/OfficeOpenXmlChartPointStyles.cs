using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native chart point-style codec; Core remains free of OpenXML dependencies.</summary>
internal static class OfficeOpenXmlChartPointStyles {
    // Preserve each format's point-override budget independently of its cache limit.
    // Count duplicate and invalid records before interpreting appearances.
    internal static IReadOnlyList<C.DataPoint> GetBoundedPoints(OpenXmlElement series, int maximum = 1_000_000) {
        List<C.DataPoint> points = series.Elements<C.DataPoint>().Take(maximum + 1).ToList();
        if (points.Count > maximum)
            throw new System.IO.InvalidDataException($"The chart exceeds the supported limit of {maximum} point overrides.");
        return points;
    }

    internal static bool IsSupported(C.ChartShapeProperties properties, A.ColorScheme? scheme, bool flattenThreeDimensional = false) {
        foreach (OpenXmlElement child in properties.ChildElements) {
            if (child is A.EffectList && !child.HasChildren && !child.HasAttributes) continue;
            if (child is A.Shape3DType && !child.HasChildren && !child.HasAttributes) continue;
            if (flattenThreeDimensional && child is A.Shape3DType) continue;
            if (child is A.NoFill) continue;
            if (child is A.SolidFill) {
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(child) ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme).HasValue) return false;
                continue;
            }
            if (child is A.PatternFill pattern) {
                if (!ReadHatch(pattern.Preset?.InnerText).HasValue ||
                    OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.GetFirstChild<A.ForegroundColor>()) ||
                    OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(pattern.GetFirstChild<A.BackgroundColor>()) ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<A.ForegroundColor>(), scheme).HasValue ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<A.BackgroundColor>(), scheme).HasValue) return false;
                continue;
            }
            if (child is A.Outline outline) {
                if (outline.Width?.Value == 0 && outline.GetFirstChild<A.NoFill>() == null) return false;
                if (outline.Width?.Value < 0 || outline.Width?.Value > 20116800 ||
                    HasUnsupportedLineAttributes(outline)) return false;
                foreach (OpenXmlElement lineChild in outline.ChildElements) {
                    if (lineChild is A.NoFill) continue;
                    if (lineChild is A.SolidFill && !OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(lineChild) &&
                        OfficeOpenXmlThemeColorResolver.ResolveColor(lineChild, scheme).HasValue) continue;
                    if (lineChild is A.Round or A.LineJoinBevel) continue;
                    if (lineChild is A.Miter miter && !miter.HasAttributes && !miter.HasChildren) continue;
                    return false;
                }
                continue;
            }
            return false;
        }
        return true;
    }

    internal static bool HasUnsupportedLineAttributes(A.Outline outline) =>
        outline.CapType != null && outline.CapType.Value != A.LineCapValues.Flat ||
        outline.Alignment != null && outline.Alignment.Value != A.PenAlignmentValues.Center ||
        outline.CompoundLineType != null && outline.CompoundLineType.Value != A.CompoundLineValues.Single;

    internal static bool HasUnsupportedPointContent(C.DataPoint point) {
        foreach (OpenXmlElement child in point.ChildElements) {
            if (child is C.Index or C.ChartShapeProperties) continue;
            if (child is C.Explosion && point.Parent is C.PieChartSeries) continue;
            if (child is C.Marker marker && !marker.HasAttributes &&
                marker.ChildElements.All(markerChild => markerChild is C.ChartShapeProperties) &&
                (marker.ChartShapeProperties == null || point.ChartShapeProperties?.HasChildren != true)) continue;
            if (child is C.InvertIfNegative invert && invert.Val?.Value == false) continue;
            if (child is C.Bubble3D bubble && bubble.Val?.Value == false) continue;
            if (child is C.ExtensionList extensions && HasOnlyUniqueIdMetadata(extensions)) continue;
            return true;
        }
        return false;
    }

    private static bool HasOnlyUniqueIdMetadata(C.ExtensionList extensions) {
        if (extensions.ChildElements.Count != 1 || extensions.FirstChild is not C.Extension extension ||
            extension.Uri?.Value != "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" ||
            extension.ChildElements.Count != 1 || extension.FirstChild is not OpenXmlElement uniqueId ||
            uniqueId.LocalName != "uniqueId" ||
            uniqueId.NamespaceUri != "http://schemas.microsoft.com/office/drawing/2014/chart" ||
            uniqueId.HasChildren) return false;
        var attributes = uniqueId.GetAttributes();
        return attributes.Count == 1 && attributes[0].LocalName == "val" &&
            attributes[0].NamespaceUri.Length == 0 && Guid.TryParse(attributes[0].Value, out _);
    }

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(OpenXmlElement series, int count, A.ColorScheme? scheme) =>
        Read(GetBoundedPoints(series), count, scheme, series.Parent?.LocalName.EndsWith("3DChart", StringComparison.Ordinal) == true);

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(IReadOnlyList<C.DataPoint> points, int count, A.ColorScheme? scheme, bool flattenThreeDimensional = false) {
        OfficeChartPointStyle?[]? styles = null;
        foreach (C.DataPoint point in points) {
            uint? index = point.Index?.Val?.Value;
            if (!index.HasValue || index.Value >= (uint)count) continue;
            if (HasUnsupportedPointContent(point))
                throw new System.IO.InvalidDataException("The native point record contains unsupported appearance or behavior.");
            C.ChartShapeProperties? properties = point.GetFirstChild<C.Marker>()?.ChartShapeProperties ??
                point.GetFirstChild<C.ChartShapeProperties>();
            if (properties == null) continue;
            if (!IsSupported(properties, scheme, flattenThreeDimensional))
                throw new System.IO.InvalidDataException("The native point appearance cannot be projected by the shared chart model.");
            OfficeColor? fill = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), scheme);
            bool noFill = properties.GetFirstChild<A.NoFill>() != null;
            A.PatternFill? pattern = properties.GetFirstChild<A.PatternFill>();
            OfficeChartHatchPattern? hatch = ReadHatch(pattern?.Preset?.InnerText);
            OfficeColor? hatchColor = hatch.HasValue ? OfficeOpenXmlThemeColorResolver.ResolveColor(pattern?.GetFirstChild<A.ForegroundColor>(), scheme) : null;
            if (hatch.HasValue && hatchColor.HasValue)
                fill = OfficeOpenXmlThemeColorResolver.ResolveColor(pattern?.GetFirstChild<A.BackgroundColor>(), scheme);
            else { hatch = null; hatchColor = null; }
            A.Outline? outline = properties.GetFirstChild<A.Outline>();
            OfficeColor? line = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme);
            bool? showOutline = outline == null ? null : outline.GetFirstChild<A.NoFill>() != null ? false : true;
            double? width = outline?.Width?.Value is int emu && emu > 0 ? emu / 12700D : null;
            OfficeStrokeLineJoin? join = outline?.GetFirstChild<A.Round>() != null ? OfficeStrokeLineJoin.Round :
                outline?.GetFirstChild<A.LineJoinBevel>() != null ? OfficeStrokeLineJoin.Bevel :
                outline?.GetFirstChild<A.Miter>() != null ? OfficeStrokeLineJoin.Miter : null;
            if (!fill.HasValue && !noFill && !hatch.HasValue && !line.HasValue && !width.HasValue && !showOutline.HasValue && !join.HasValue) continue;
            styles ??= new OfficeChartPointStyle?[count];
            styles[(int)index.Value] = new OfficeChartPointStyle(fill, noFill, hatch, hatchColor, line, width, showOutline, join);
        }
        return styles;
    }

    internal static void ApplySeries(OpenXmlCompositeElement series, OfficeChartSeries data) {
        if (data.PointStyles == null) return;
        ApplyPoints(series, data.PointStyles, data.PointColors, data.Values.Count);
    }

    internal static void ApplyPoints(OpenXmlCompositeElement series,
        IReadOnlyList<OfficeChartPointStyle?>? styles, IReadOnlyList<OfficeColor?>? colors, int count) {
        if (styles == null && colors == null) return;
        var existing = new Dictionary<uint, C.DataPoint>();
        foreach (C.DataPoint point in GetBoundedPoints(series)) {
            if (point.Index?.Val?.Value is uint index && !existing.ContainsKey(index))
                existing.Add(index, point);
        }
        OpenXmlElement? anchor = FindPointAnchor(series);
        for (int index = 0; index < count; index++) {
            OfficeChartPointStyle? style = styles?[index];
            OfficeColor? color = colors?[index];
            existing.TryGetValue((uint)index, out C.DataPoint? point);
            if (point == null && style == null && !color.HasValue) continue;
            ApplyPointCore(series, (uint)index, style, color, point, anchor);
        }
    }

    internal static void ApplyPoint(OpenXmlCompositeElement series, uint index, OfficeChartPointStyle? style, OfficeColor? legacyColor = null) {
        if (style?.OutlineWidth > 1584) throw new ArgumentOutOfRangeException(nameof(style), "Native chart outlines cannot exceed 1584 points.");
        if (!SupportsDataPoints(series))
            throw new NotSupportedException("Point styles are not supported by this native chart series.");
        C.DataPoint? point = series.Elements<C.DataPoint>().FirstOrDefault(item => item.Index?.Val?.Value == index);
        if (point == null && style == null && !legacyColor.HasValue) return;
        ApplyPointCore(series, index, style, legacyColor, point, FindPointAnchor(series));
    }

    private static OpenXmlElement? FindPointAnchor(OpenXmlCompositeElement series) =>
        series.ChildElements.FirstOrDefault(child =>
            child is C.DataLabels or C.Trendline or C.ErrorBars or C.CategoryAxisData or C.Values or
                C.XValues or C.YValues or C.BubbleSize or C.Smooth or C.ExtensionList);

    private static bool SupportsDataPoints(OpenXmlCompositeElement series) =>
        series is C.BarChartSeries or C.LineChartSeries or C.AreaChartSeries or C.PieChartSeries or
            C.ScatterChartSeries or C.RadarChartSeries or C.BubbleChartSeries;

    private static void ApplyPointCore(OpenXmlCompositeElement series, uint index,
        OfficeChartPointStyle? style, OfficeColor? legacyColor, C.DataPoint? point, OpenXmlElement? anchor) {
        if (style?.OutlineWidth > 1584) throw new ArgumentOutOfRangeException(nameof(style), "Native chart outlines cannot exceed 1584 points.");
        if (point == null) {
            point = new C.DataPoint(new C.Index { Val = index });
            if (anchor != null) series.InsertBefore(point, anchor);
            else series.Append(point);
        }
        ApplyPointStyle(point, style, legacyColor);
    }

    private static void ApplyPointStyle(C.DataPoint point, OfficeChartPointStyle? style, OfficeColor? legacyColor) {
        C.Marker? pointMarker = point.GetFirstChild<C.Marker>();
        OpenXmlCompositeElement styleOwner = pointMarker?.GetFirstChild<C.ChartShapeProperties>() != null
            ? pointMarker : point;
        C.ChartShapeProperties properties = styleOwner.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.RemoveAllChildren<A.Outline>();
        if (style?.NoFill == true) properties.AddChild(new A.NoFill(), true);
        else if (style?.Hatch is OfficeChartHatchPattern hatch) {
            var preset = new EnumValue<A.PresetPatternValues> { InnerText = HatchToken(hatch) };
            properties.AddChild(new A.PatternFill(
                new A.ForegroundColor(CreateColor(style.HatchColor!.Value)),
                new A.BackgroundColor(CreateColor(style.FillColor ?? legacyColor ?? OfficeColor.White))) { Preset = preset }, true);
        } else if ((style?.FillColor ?? legacyColor) is OfficeColor fill)
            properties.AddChild(new A.SolidFill(CreateColor(fill)), true);
        if (style?.ShowOutline == false) properties.AddChild(new A.Outline(new A.NoFill()), true);
        else if (style?.OutlineColor != null || style?.OutlineWidth != null || style?.ShowOutline == true || style?.OutlineJoin != null) {
            var outline = new A.Outline();
            if (style.OutlineWidth.HasValue) outline.Width = (int)Math.Round(style.OutlineWidth.Value * 12700);
            if (style.OutlineColor is OfficeColor line) outline.Append(new A.SolidFill(CreateColor(line)));
            if (style.OutlineJoin is OfficeStrokeLineJoin join) outline.Append(join switch {
                OfficeStrokeLineJoin.Round => new A.Round(),
                OfficeStrokeLineJoin.Bevel => new A.LineJoinBevel(),
                _ => new A.Miter()
            });
            properties.AddChild(outline, true);
        }
        if (properties.Parent == null && properties.HasChildren) styleOwner.AddChild(properties, true);
        else if (!properties.HasChildren && properties.Parent != null) properties.Remove();
    }

    private static A.RgbColorModelHex CreateColor(OfficeColor color) {
        var value = new A.RgbColorModelHex { Val = color.ToRgbHex() };
        if (color.A != 255) value.Append(new A.Alpha { Val = (int)Math.Round(color.A * 100000D / 255) });
        return value;
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
}
