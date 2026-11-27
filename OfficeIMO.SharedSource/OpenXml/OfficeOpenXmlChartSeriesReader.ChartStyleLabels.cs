using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    /// <summary>The most characters a chart style part may hold for its data-label style to be read.</summary>
    internal const int MaximumChartStyleCharacters = 131072;

    private static NativeText ReadChartStyleLabelText(IOfficeChartPart chartPart, OfficeXmlNode scheme) {
        const string chartStyleNamespace = ChartStyleUri;
        const string drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        IReadOnlyList<Func<Stream>> parts = chartPart.ChartStyles(2);
        if (parts.Count == 0) return default;
        if (parts.Count != 1) throw new NotSupportedException("Multiple chart text styles cannot be projected.");
        try {
            // The limit applies while the part streams in, so an oversized style is refused before it is parsed.
            using Stream stream = parts[0]();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumChartStyleCharacters
            });
            XElement? root = XDocument.Load(reader).Root;
            if (root?.Name != XName.Get("chartStyle", chartStyleNamespace))
                throw new NotSupportedException("The chart style cannot be projected.");
            XElement[] rules = root.Elements(XName.Get("dataLabel", chartStyleNamespace)).Take(2).ToArray();
            if (rules.Length == 0) return default;
            if (rules.Length != 1) throw new NotSupportedException("Multiple chart data-label styles cannot be projected.");
            XElement rule = rules[0];
            if (rule.HasAttributes || rule.Elements().Any(child =>
                    child.Name != XName.Get("lnRef", chartStyleNamespace) &&
                    child.Name != XName.Get("fillRef", chartStyleNamespace) &&
                    child.Name != XName.Get("effectRef", chartStyleNamespace) &&
                    child.Name != XName.Get("fontRef", chartStyleNamespace) &&
                    child.Name != XName.Get("defRPr", chartStyleNamespace)))
                throw new NotSupportedException("The chart data-label appearance cannot be projected.");
            if (rule.Element(XName.Get("spPr", chartStyleNamespace)) != null ||
                rule.Element(XName.Get("lnRef", chartStyleNamespace))?.Attribute("idx")?.Value is string lineIndex && lineIndex != "0" ||
                rule.Element(XName.Get("fillRef", chartStyleNamespace))?.Attribute("idx")?.Value is string fillIndex && fillIndex != "0" ||
                rule.Element(XName.Get("effectRef", chartStyleNamespace))?.Attribute("idx")?.Value is string effectIndex && effectIndex != "0")
                throw new NotSupportedException("The chart data-label shape style cannot be projected.");

            XElement? font = rule.Element(XName.Get("fontRef", chartStyleNamespace));
            string? fontReference = font?.Attribute("idx")?.Value;
            if (fontReference != null && fontReference is not "minor" and not "major")
                throw new NotSupportedException("The chart data-label font reference cannot be projected.");
            OfficeXmlNode fontScheme = fontReference == null ? default : chartPart.HostFontScheme;
            string? family = fontReference == "major" ? SchemeTypeface(fontScheme, "majorFont") :
                fontReference == "minor" ? SchemeTypeface(fontScheme, "minorFont") : null;
            if (fontReference != null && string.IsNullOrWhiteSpace(family))
                throw new NotSupportedException("The chart data-label font family cannot be resolved.");

            OfficeColor? color = null;
            XElement[] colorElements = font?.Elements().Take(2).ToArray() ?? Array.Empty<XElement>();
            if (colorElements.Length > 1) throw new NotSupportedException("Multiple chart data-label font colors cannot be projected.");
            if (colorElements.Length == 1) {
                XElement element = colorElements[0];
                if (element.Name.NamespaceName != drawingNamespace)
                    throw new NotSupportedException("The chart data-label font color cannot be projected.");
                if (element.Name.LocalName is not "schemeClr" and not "srgbClr")
                    throw new NotSupportedException("The chart data-label font color cannot be projected.");
                // The color element on its own, with its namespace declared, as the color resolver reads it.
                byte[] colorMarkup = System.Text.Encoding.UTF8.GetBytes(element.ToString(SaveOptions.DisableFormatting));
                OfficeXmlTree colorTree = OfficeXmlTree.Parse(colorMarkup, colorMarkup.Length, new OfficeXmlNamespaces.Table());
                OfficeXmlNode nativeColor = new(colorTree, colorTree.Root);
                // A scheme color goes through the chart's color map, as the slot it names is mapped before it resolves.
                OfficeXmlNode colorMap = element.Name.LocalName == "schemeClr" && element.Attribute("val") != null
                    ? OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart) : default;
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(nativeColor))
                    throw new NotSupportedException("The chart data-label font transforms cannot be projected.");
                color = OfficeOpenXmlThemeColorResolver.ResolveColor(nativeColor, scheme, colorMap: colorMap) ??
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
        } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
            throw new NotSupportedException("The chart data-label style XML cannot be projected.", exception);
        }
    }
}
