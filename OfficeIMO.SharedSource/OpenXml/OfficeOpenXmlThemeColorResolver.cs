using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>
/// Resolves DrawingML theme colors and applies their ordered color transformations.
/// This source is shared by the Word, Excel, and PowerPoint OpenXML adapters.
/// </summary>
internal static class OfficeOpenXmlThemeColorResolver {
    private static readonly string[] DefaultSpreadsheetThemeColors = {
        "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D",
        "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"
    };

    /// <summary>Reads a bounded radial palette from a modern color style or classic style 2.</summary>
    internal static OfficeColor[]? ReadRadialPalette(ChartPart chartPart, OpenXmlCompositeElement series,
        int pointCount, A.ColorScheme? scheme) {
        if (scheme == null || pointCount < 1 || pointCount > 6 ||
            series.Parent is not C.PieChart and not C.DoughnutChart ||
            series.Parent.GetFirstChild<C.VaryColors>() is not C.VaryColors varyColors ||
            varyColors.Val?.Value == false)
            return null;
        // Point fills take precedence over the theme palette in every static renderer.
        // An unsupported inherited palette is immaterial when every rendered ring has
        // an explicit appearance for every category.
        if (AllRadialPointsHaveExplicitFill((OpenXmlCompositeElement)series.Parent, pointCount)) return null;
        ChartColorStylePart? colorStylePart = chartPart.GetPartsOfType<ChartColorStylePart>().FirstOrDefault();
        if (colorStylePart != null) {
            ChartStylePart? stylePart = chartPart.GetPartsOfType<ChartStylePart>().FirstOrDefault();
            if (stylePart == null || !HasAutomaticDataPointFill(stylePart))
                throw new NotSupportedException("The modern chart style does not use the automatic data-point palette.");
            bool multipleRings = series.Parent.Elements<C.PieChartSeries>().Skip(1).Any();
            return ReadModernRadialPalette(colorStylePart, pointCount, scheme, multipleRings)
                ?? throw new NotSupportedException("The modern radial color style cannot be projected.");
        }
        if (chartPart.GetPartsOfType<ChartStylePart>().Any()) return null;
        C.Style? style = series.Ancestors<C.ChartSpace>().FirstOrDefault()?
            .Descendants<C.Style>().FirstOrDefault();
        if (style != null && style.Val?.Value != 2)
            return null;
        var colors = new OfficeColor[pointCount];
        for (int index = 0; index < pointCount; index++) {
            OfficeColor? color = ResolveSchemeColor(scheme, "accent" + (index + 1));
            if (!color.HasValue) return null;
            colors[index] = color.Value;
        }
        return colors;
    }

    private static bool AllRadialPointsHaveExplicitFill(OpenXmlCompositeElement chart, int pointCount) {
        bool hasSeries = false;
        foreach (C.PieChartSeries series in chart.Elements<C.PieChartSeries>()) {
            hasSeries = true;
            var covered = new bool[pointCount];
            foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
                uint? index = point.Index?.Val?.Value;
                if (!index.HasValue || index.Value >= (uint)pointCount) continue;
                C.ChartShapeProperties? appearance = point.GetFirstChild<C.ChartShapeProperties>();
                if (appearance?.GetFirstChild<A.SolidFill>() != null ||
                    appearance?.GetFirstChild<A.NoFill>() != null ||
                    appearance?.GetFirstChild<A.PatternFill>() != null)
                    covered[(int)index.Value] = true;
            }
            if (covered.Any(value => !value)) return false;
        }
        return hasSeries;
    }

    private static bool HasAutomaticDataPointFill(ChartStylePart part) {
        const string chartStyleNamespace = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        try {
            using var stream = part.GetStream();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 131072
            });
            XElement? root = XDocument.Load(reader).Root;
            XElement? dataPoint = root?.Element(XName.Get("dataPoint", chartStyleNamespace));
            XElement? fill = dataPoint?.Element(XName.Get("fillRef", chartStyleNamespace));
            XElement? color = fill?.Element(XName.Get("styleClr", chartStyleNamespace));
            return root?.Name == XName.Get("chartStyle", chartStyleNamespace) &&
                dataPoint != null && dataPoint.Element(XName.Get("spPr", chartStyleNamespace)) == null &&
                fill?.Attribute("idx")?.Value == "1" &&
                fill.Elements().Count() == 1 && color?.Attribute("val")?.Value == "auto" &&
                !color.HasElements;
        } catch (XmlException) {
            return false;
        }
    }

    private static OfficeColor[]? ReadModernRadialPalette(ChartColorStylePart part,
        int pointCount, A.ColorScheme scheme, bool multipleRings) {
        const string chartStyleNamespace = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        const string drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        try {
            using var stream = part.GetStream();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 131072
            });
            XElement? root = XDocument.Load(reader).Root;
            if (root?.Name != XName.Get("colorStyle", chartStyleNamespace) ||
                (string?)root.Attribute("meth") != "cycle") return null;
            XElement? firstVariation = root.Elements(XName.Get("variation", chartStyleNamespace)).FirstOrDefault();
            if (firstVariation is { HasElements: true }) return null;
            if (multipleRings && root.Elements(XName.Get("variation", chartStyleNamespace))
                .Any(variation => variation.HasElements)) return null;
            XElement[] entries = root.Elements().Where(element => element.Name.NamespaceName == drawingNamespace)
                .Take(pointCount + 1).ToArray();
            if (entries.Length < pointCount) return null;
            var colors = new OfficeColor[pointCount];
            for (int index = 0; index < pointCount; index++) {
                XElement entry = entries[index];
                if (entry.Name != XName.Get("schemeClr", drawingNamespace) || entry.HasElements) return null;
                OfficeColor? color = ResolveSchemeColor(scheme, (string?)entry.Attribute("val"));
                if (!color.HasValue) return null;
                colors[index] = color.Value;
            }
            return colors;
        } catch (XmlException) {
            return null;
        }
    }

    internal static OfficeColor? ResolveColor(
        OpenXmlElement? container,
        A.ColorScheme? colorScheme,
        OpenXmlElement? placeholderColor = null) {
        OpenXmlElement? colorElement = FindColorElement(container);
        return ResolveColorElement(colorElement, colorScheme,
            FindColorElement(placeholderColor));
    }

    internal static bool HasUnsupportedTransforms(OpenXmlElement? container) {
        OpenXmlElement? color = FindColorElement(container);
        if (color == null) return false;
        foreach (OpenXmlElement transform in color.ChildElements) {
            if (transform.LocalName is "comp" or "inv" or "gray") continue;
            if (transform.LocalName is not ("alpha" or "alphaMod" or "alphaOff" or "tint" or "shade" or
                "lumMod" or "lumOff" or "red" or "redMod" or "redOff" or "green" or "greenMod" or
                "greenOff" or "blue" or "blueMod" or "blueOff") || !TryReadTransformValue(transform, out _)) return true;
        }
        return false;
    }

    internal static OfficeColor? ResolveSchemeColor(A.ColorScheme? colorScheme, string? scheme) {
        if (colorScheme == null || string.IsNullOrWhiteSpace(scheme)) {
            return null;
        }

        string normalized = scheme!.Trim().ToLowerInvariant();
        OpenXmlCompositeElement? colorElement = normalized switch {
            "dark1" or "dk1" or "text1" or "tx1" => colorScheme.GetFirstChild<A.Dark1Color>(),
            "light1" or "lt1" or "background1" or "bg1" => colorScheme.GetFirstChild<A.Light1Color>(),
            "dark2" or "dk2" or "text2" or "tx2" => colorScheme.GetFirstChild<A.Dark2Color>(),
            "light2" or "lt2" or "background2" or "bg2" => colorScheme.GetFirstChild<A.Light2Color>(),
            "accent1" => colorScheme.GetFirstChild<A.Accent1Color>(),
            "accent2" => colorScheme.GetFirstChild<A.Accent2Color>(),
            "accent3" => colorScheme.GetFirstChild<A.Accent3Color>(),
            "accent4" => colorScheme.GetFirstChild<A.Accent4Color>(),
            "accent5" => colorScheme.GetFirstChild<A.Accent5Color>(),
            "accent6" => colorScheme.GetFirstChild<A.Accent6Color>(),
            "hyperlink" or "hlink" => colorScheme.GetFirstChild<A.Hyperlink>(),
            "followedhyperlink" or "folhlink" => colorScheme.GetFirstChild<A.FollowedHyperlinkColor>(),
            _ => null
        };

        return ResolveThemeEntry(colorElement);
    }

    internal static OfficeColor? ResolveSpreadsheetThemeColor(A.ColorScheme? colorScheme, uint themeIndex) {
        string? scheme = themeIndex switch {
            0U => "light1",
            1U => "dark1",
            2U => "light2",
            3U => "dark2",
            4U => "accent1",
            5U => "accent2",
            6U => "accent3",
            7U => "accent4",
            8U => "accent5",
            9U => "accent6",
            10U => "hyperlink",
            11U => "followedHyperlink",
            _ => null
        };
        OfficeColor? resolved = ResolveSchemeColor(colorScheme, scheme);
        if (resolved.HasValue || themeIndex >= DefaultSpreadsheetThemeColors.Length) {
            return resolved;
        }

        return OfficeColor.TryParseHex(DefaultSpreadsheetThemeColors[themeIndex], out OfficeColor fallback)
            ? fallback
            : (OfficeColor?)null;
    }

    private static OfficeColor? ResolveColorElement(
        OpenXmlElement? colorElement,
        A.ColorScheme? colorScheme,
        OpenXmlElement? placeholderColor) {
        if (colorElement == null) {
            return null;
        }

        OfficeColor? color;
        if (colorElement is A.RgbColorModelHex rgbColor) {
            color = ParseRgb(rgbColor.Val?.Value);
        } else if (colorElement is A.RgbColorModelPercentage scRgbColor) {
            color = ParseScRgb(scRgbColor);
        } else if (colorElement is A.HslColor hslColor) {
            color = ParseHsl(hslColor);
        } else if (colorElement is A.SystemColor systemColor) {
            color = ParseRgb(systemColor.LastColor?.Value);
        } else if (colorElement is A.SchemeColor schemeColor) {
            string? scheme = GetSchemeValue(schemeColor);
            if (IsPlaceholderScheme(scheme)) {
                color = ResolveColorElement(placeholderColor, colorScheme, null);
            } else {
                color = ResolveSchemeColor(colorScheme, scheme);
            }
        } else if (colorElement is A.PresetColor presetColor) {
            color = OfficeColor.TryParse(presetColor.Val?.Value.ToString(), out OfficeColor preset)
                ? preset
                : (OfficeColor?)null;
        } else {
            color = null;
        }

        return color.HasValue ? ApplyTransforms(color.Value, colorElement) : null;
    }

    private static OfficeColor ApplyTransforms(OfficeColor color, OpenXmlElement colorElement) {
        OfficeColor resolved = color;
        foreach (OpenXmlElement transform in colorElement.ChildElements) {
            switch (transform.LocalName) {
                case "comp":
                    resolved = OfficeColorTransforms.Complement(resolved);
                    continue;
                case "inv":
                    resolved = OfficeColor.FromRgba(
                        (byte)(255 - resolved.R),
                        (byte)(255 - resolved.G),
                        (byte)(255 - resolved.B),
                        resolved.A);
                    continue;
                case "gray":
                    byte gray = ToChannel((resolved.R * 0.299D) + (resolved.G * 0.587D) + (resolved.B * 0.114D));
                    resolved = OfficeColor.FromRgba(gray, gray, gray, resolved.A);
                    continue;
            }

            if (!TryReadTransformValue(transform, out double value)) {
                continue;
            }

            switch (transform.LocalName) {
                case "alpha":
                    resolved = OfficeColorTransforms.WithAlpha(resolved, ClampUnit(value));
                    break;
                case "alphaMod":
                    resolved = OfficeColorTransforms.ModulateAlpha(resolved, Math.Max(0D, value));
                    break;
                case "alphaOff":
                    resolved = OfficeColorTransforms.OffsetAlpha(resolved, value);
                    break;
                case "tint":
                    resolved = OfficeColorTransforms.Tint(resolved, ClampUnit(value));
                    break;
                case "shade":
                    resolved = OfficeColorTransforms.Shade(resolved, ClampUnit(value));
                    break;
                case "lumMod":
                    resolved = OfficeColorTransforms.ModulateLuminance(resolved, Math.Max(0D, value));
                    break;
                case "lumOff":
                    resolved = OfficeColorTransforms.OffsetLuminance(resolved, value);
                    break;
                case "red":
                    resolved = OfficeColor.FromRgba(ToChannel(255D * value), resolved.G, resolved.B, resolved.A);
                    break;
                case "redMod":
                    resolved = OfficeColor.FromRgba(ToChannel(resolved.R * value), resolved.G, resolved.B, resolved.A);
                    break;
                case "redOff":
                    resolved = OfficeColor.FromRgba(ToChannel(resolved.R + (255D * value)), resolved.G, resolved.B, resolved.A);
                    break;
                case "green":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(255D * value), resolved.B, resolved.A);
                    break;
                case "greenMod":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G * value), resolved.B, resolved.A);
                    break;
                case "greenOff":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G + (255D * value)), resolved.B, resolved.A);
                    break;
                case "blue":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(255D * value), resolved.A);
                    break;
                case "blueMod":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B * value), resolved.A);
                    break;
                case "blueOff":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B + (255D * value)), resolved.A);
                    break;
            }
        }

        return resolved;
    }

    private static OpenXmlElement? FindColorElement(OpenXmlElement? container) {
        if (container == null) {
            return null;
        }

        if (container is A.RgbColorModelHex or A.RgbColorModelPercentage
            or A.HslColor or A.SystemColor or A.SchemeColor or A.PresetColor) {
            return container;
        }

        return container.GetFirstChild<A.RgbColorModelHex>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.RgbColorModelPercentage>()
            ?? container.GetFirstChild<A.HslColor>()
            ?? container.GetFirstChild<A.SchemeColor>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.SystemColor>()
            ?? container.GetFirstChild<A.PresetColor>();
    }

    private static OfficeColor? ResolveThemeEntry(OpenXmlCompositeElement? colorElement) {
        if (colorElement == null) {
            return null;
        }

        return ResolveColorElement(FindColorElement(colorElement),
            colorScheme: null, placeholderColor: null);
    }

    private static OfficeColor? ParseRgb(string? value) =>
        OfficeColor.TryParseHex(value, out OfficeColor color) ? color : (OfficeColor?)null;

    private static OfficeColor? ParseScRgb(A.RgbColorModelPercentage color) {
        if (color.RedPortion?.Value is not int red
            || color.GreenPortion?.Value is not int green
            || color.BluePortion?.Value is not int blue) {
            return null;
        }
        return OfficeColor.FromRgb(
            ToChannel(255D * LinearScRgbToSrgb(red / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(green / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(blue / 100000D)));
    }

    private static double LinearScRgbToSrgb(double value) {
        double linear = ClampUnit(value);
        return linear <= 0.0031308D
            ? 12.92D * linear
            : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D;
    }

    private static OfficeColor? ParseHsl(A.HslColor color) {
        if (color.HueValue?.Value is not int hue
            || color.SatValue?.Value is not int saturation
            || color.LumValue?.Value is not int luminance) {
            return null;
        }
        double normalizedHue = ((hue / 60000D) % 360D + 360D) % 360D;
        double normalizedSaturation = ClampUnit(saturation / 100000D);
        double normalizedLuminance = ClampUnit(luminance / 100000D);
        double chroma = (1D - Math.Abs(2D * normalizedLuminance - 1D))
            * normalizedSaturation;
        double sector = normalizedHue / 60D;
        double intermediate = chroma * (1D - Math.Abs(sector % 2D - 1D));
        double red;
        double green;
        double blue;
        if (sector < 1D) {
            red = chroma; green = intermediate; blue = 0D;
        } else if (sector < 2D) {
            red = intermediate; green = chroma; blue = 0D;
        } else if (sector < 3D) {
            red = 0D; green = chroma; blue = intermediate;
        } else if (sector < 4D) {
            red = 0D; green = intermediate; blue = chroma;
        } else if (sector < 5D) {
            red = intermediate; green = 0D; blue = chroma;
        } else {
            red = chroma; green = 0D; blue = intermediate;
        }
        double match = normalizedLuminance - chroma / 2D;
        return OfficeColor.FromRgb(ToChannel(255D * (red + match)),
            ToChannel(255D * (green + match)),
            ToChannel(255D * (blue + match)));
    }

    private static string? GetSchemeValue(A.SchemeColor? schemeColor) {
        string? attribute = schemeColor?.GetAttribute("val", string.Empty).Value;
        return !string.IsNullOrWhiteSpace(attribute)
            ? attribute
            : schemeColor?.Val?.Value.ToString();
    }

    private static bool IsPlaceholderScheme(string? scheme) =>
        string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadTransformValue(OpenXmlElement transform, out double value) {
        value = 0D;
        string? raw = transform.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(raw) || !int.TryParse(raw, out int scaled)) {
            return false;
        }

        value = scaled / 100000D;
        return true;
    }

    private static byte ToChannel(double value) =>
        (byte)Math.Max(0D, Math.Min(255D, Math.Round(value, MidpointRounding.ToEven)));

    private static double ClampUnit(double value) => Math.Max(0D, Math.Min(1D, value));
}
