using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    private static NativeText ReadChartStyleLabelText(ChartPart chartPart, A.ColorScheme? scheme) {
        ChartStylePart[] parts = chartPart.GetPartsOfType<ChartStylePart>().Take(2).ToArray();
        if (parts.Length == 0) return default;
        if (parts.Length != 1) throw new NotSupportedException("Multiple chart text styles cannot be projected.");
        const string chartStyleNamespace = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        const string drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        try {
            using Stream stream = parts[0].GetStream();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 131072
            });
            XElement? root = XDocument.Load(reader).Root;
            if (root?.Name != XName.Get("chartStyle", chartStyleNamespace))
                throw new NotSupportedException("The chart style cannot be projected.");
            XElement[] rules = root.Elements(XName.Get("dataLabel", chartStyleNamespace)).Take(2).ToArray();
            if (rules.Length == 0) return default;
            if (rules.Length != 1) throw new NotSupportedException("Multiple chart data-label styles cannot be projected.");
            XElement rule = rules[0];
            if (rule.Element(XName.Get("spPr", chartStyleNamespace)) != null ||
                rule.Element(XName.Get("lnRef", chartStyleNamespace))?.Attribute("idx")?.Value is string lineIndex && lineIndex != "0" ||
                rule.Element(XName.Get("fillRef", chartStyleNamespace))?.Attribute("idx")?.Value is string fillIndex && fillIndex != "0" ||
                rule.Element(XName.Get("effectRef", chartStyleNamespace))?.Attribute("idx")?.Value is string effectIndex && effectIndex != "0")
                throw new NotSupportedException("The chart data-label shape style cannot be projected.");

            XElement? font = rule.Element(XName.Get("fontRef", chartStyleNamespace));
            string? fontReference = font?.Attribute("idx")?.Value;
            if (fontReference != null && fontReference is not "minor" and not "major")
                throw new NotSupportedException("The chart data-label font reference cannot be projected.");
            A.FontScheme? fontScheme = fontReference == null ? null : ResolveChartFontScheme(chartPart);
            string? family = fontReference == "major" ? fontScheme?.MajorFont?.LatinFont?.Typeface?.Value :
                fontReference == "minor" ? fontScheme?.MinorFont?.LatinFont?.Typeface?.Value : null;
            if (fontReference != null && string.IsNullOrWhiteSpace(family))
                throw new NotSupportedException("The chart data-label font family cannot be resolved.");

            OfficeColor? color = null;
            XElement[] colorElements = font?.Elements().Take(2).ToArray() ?? Array.Empty<XElement>();
            if (colorElements.Length > 1) throw new NotSupportedException("Multiple chart data-label font colors cannot be projected.");
            if (colorElements.Length == 1) {
                XElement element = colorElements[0];
                if (element.Name.NamespaceName != drawingNamespace)
                    throw new NotSupportedException("The chart data-label font color cannot be projected.");
                DocumentFormat.OpenXml.OpenXmlElement nativeColor = element.Name.LocalName switch {
                    "schemeClr" => new A.SchemeColor(element.ToString(SaveOptions.DisableFormatting)),
                    "srgbClr" => new A.RgbColorModelHex(element.ToString(SaveOptions.DisableFormatting)),
                    _ => throw new NotSupportedException("The chart data-label font color cannot be projected.")
                };
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(nativeColor))
                    throw new NotSupportedException("The chart data-label font transforms cannot be projected.");
                color = OfficeOpenXmlThemeColorResolver.ResolveColor(nativeColor, scheme) ??
                    throw new NotSupportedException("The chart data-label font color cannot be resolved.");
            }

            XElement? run = rule.Element(XName.Get("defRPr", chartStyleNamespace));
            double? size = null;
            OfficeFontStyle? style = null;
            if (run != null) {
                foreach (XAttribute attribute in run.Attributes()) {
                    switch (attribute.Name.LocalName) {
                        case "sz":
                            if (!int.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int hundredths) || hundredths <= 0)
                                throw new NotSupportedException("The chart data-label font size cannot be projected.");
                            size = hundredths / 100d;
                            break;
                        case "b":
                            if (attribute.Value is "1" or "true") style = (style ?? OfficeFontStyle.Regular) | OfficeFontStyle.Bold;
                            else if (attribute.Value is not "0" and not "false") throw new NotSupportedException("The chart data-label bold style cannot be projected.");
                            break;
                        case "i":
                            if (attribute.Value is "1" or "true") style = (style ?? OfficeFontStyle.Regular) | OfficeFontStyle.Italic;
                            else if (attribute.Value is not "0" and not "false") throw new NotSupportedException("The chart data-label italic style cannot be projected.");
                            break;
                        case "kern":
                            break;
                        default:
                            throw new NotSupportedException("The chart data-label run style cannot be projected.");
                    }
                }
                if (run.HasElements) throw new NotSupportedException("The chart data-label run effects cannot be projected.");
            }
            return new NativeText(size, style, color, family);
        } catch (XmlException exception) {
            throw new NotSupportedException("The chart data-label style XML cannot be projected.", exception);
        }
    }

    private static A.FontScheme? ResolveChartFontScheme(ChartPart chartPart) {
        var visited = new System.Collections.Generic.HashSet<OpenXmlPart>();
        var queue = new System.Collections.Generic.Queue<OpenXmlPart>(chartPart.GetParentParts());
        while (queue.Count > 0 && visited.Count < 16) {
            OpenXmlPart owner = queue.Dequeue();
            if (!visited.Add(owner)) continue;
            A.FontScheme? scheme = owner switch {
                SlidePart slide => slide.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                    slide.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                    slide.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                SlideLayoutPart layout => layout.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                    layout.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                SlideMasterPart master => master.ThemePart?.Theme?.ThemeElements?.FontScheme,
                NotesSlidePart notes => notes.ThemeOverridePart?.ThemeOverride?.FontScheme ??
                    notes.NotesMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme,
                NotesMasterPart notesMaster => notesMaster.ThemePart?.Theme?.ThemeElements?.FontScheme,
                HandoutMasterPart handout => handout.ThemePart?.Theme?.ThemeElements?.FontScheme,
                MainDocumentPart document => document.ThemePart?.Theme?.ThemeElements?.FontScheme,
                WorkbookPart workbook => workbook.ThemePart?.Theme?.ThemeElements?.FontScheme,
                _ => null
            };
            if (scheme != null) return scheme;
            foreach (OpenXmlPart parent in owner.GetParentParts()) queue.Enqueue(parent);
        }
        return null;
    }
}
