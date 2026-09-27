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
    internal static bool IsSupported(C.ChartShapeProperties properties, A.ColorScheme? scheme) {
        foreach (OpenXmlElement child in properties.ChildElements) {
            if (child is A.NoFill) continue;
            if (child is A.SolidFill) {
                if (!OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme).HasValue) return false;
                continue;
            }
            if (child is A.PatternFill pattern) {
                if (!ReadHatch(pattern.Preset?.InnerText).HasValue ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<A.ForegroundColor>(), scheme).HasValue ||
                    !OfficeOpenXmlThemeColorResolver.ResolveColor(pattern.GetFirstChild<A.BackgroundColor>(), scheme).HasValue) return false;
                continue;
            }
            if (child is A.Outline outline) {
                if (outline.Width?.Value < 0 || outline.Width?.Value > 20116800 ||
                    outline.CapType != null || outline.Alignment != null || outline.CompoundLineType != null) return false;
                foreach (OpenXmlElement lineChild in outline.ChildElements) {
                    if (lineChild is A.NoFill) continue;
                    if (lineChild is A.SolidFill && OfficeOpenXmlThemeColorResolver.ResolveColor(lineChild, scheme).HasValue) continue;
                    return false;
                }
                continue;
            }
            return false;
        }
        return true;
    }

    internal static IReadOnlyList<OfficeChartPointStyle?>? Read(OpenXmlElement series, int count, A.ColorScheme? scheme) {
        OfficeChartPointStyle?[]? styles = null;
        foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
            uint? index = point.Index?.Val?.Value;
            if (!index.HasValue || index.Value >= (uint)count) continue;
            C.ChartShapeProperties? properties = point.GetFirstChild<C.ChartShapeProperties>();
            if (properties == null) continue;
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
            bool? showOutline = outline?.GetFirstChild<A.NoFill>() != null ? false : line.HasValue ? true : null;
            double? width = outline?.Width?.Value is int emu && emu > 0 ? emu / 12700D : null;
            if (!fill.HasValue && !noFill && !hatch.HasValue && !line.HasValue && !width.HasValue && !showOutline.HasValue) continue;
            styles ??= new OfficeChartPointStyle?[count];
            styles[(int)index.Value] = new OfficeChartPointStyle(fill, noFill, hatch, hatchColor, line, width, showOutline);
        }
        return styles;
    }

    internal static void ApplySeries(OpenXmlCompositeElement series, OfficeChartSeries data) {
        if (data.PointStyles == null) return;
        for (int index = 0; index < data.PointStyles.Count; index++)
            ApplyPoint(series, (uint)index, data.PointStyles[index], data.PointColors?[index]);
    }

    internal static void ApplyPoint(OpenXmlCompositeElement series, uint index, OfficeChartPointStyle? style, OfficeColor? legacyColor = null) {
        if (style?.OutlineWidth > 1584) throw new ArgumentOutOfRangeException(nameof(style), "Native chart outlines cannot exceed 1584 points.");
        C.DataPoint? point = series.Elements<C.DataPoint>().FirstOrDefault(item => item.Index?.Val?.Value == index);
        if (point == null) {
            if (style == null && !legacyColor.HasValue) return;
            point = new C.DataPoint(new C.Index { Val = index });
            OpenXmlElement? anchor = series.ChildElements.FirstOrDefault(child =>
                child is C.DataLabels or C.Trendline or C.ErrorBars or C.CategoryAxisData or C.Values or
                    C.XValues or C.YValues or C.BubbleSize or C.Smooth or C.ExtensionList);
            if (anchor != null) series.InsertBefore(point, anchor);
            else series.Append(point);
        }
        C.ChartShapeProperties properties = point.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
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
        else if (style?.OutlineColor != null || style?.OutlineWidth != null || style?.ShowOutline == true) {
            var outline = new A.Outline();
            if (style.OutlineWidth.HasValue) outline.Width = (int)Math.Round(style.OutlineWidth.Value * 12700);
            if (style.OutlineColor is OfficeColor line) outline.Append(new A.SolidFill(CreateColor(line)));
            properties.AddChild(outline, true);
        }
        if (properties.Parent == null && properties.HasChildren) point.AddChild(properties, true);
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
        "cross" => OfficeChartHatchPattern.Cross, "diagCross" => OfficeChartHatchPattern.DiagonalCross, _ => null
    };
    private static string HatchToken(OfficeChartHatchPattern hatch) => hatch switch {
        OfficeChartHatchPattern.Horizontal => "horz", OfficeChartHatchPattern.Vertical => "vert",
        OfficeChartHatchPattern.ForwardDiagonal => "upDiag", OfficeChartHatchPattern.BackwardDiagonal => "dnDiag",
        OfficeChartHatchPattern.Cross => "cross", OfficeChartHatchPattern.DiagonalCross => "diagCross",
        _ => throw new ArgumentOutOfRangeException(nameof(hatch))
    };
}
