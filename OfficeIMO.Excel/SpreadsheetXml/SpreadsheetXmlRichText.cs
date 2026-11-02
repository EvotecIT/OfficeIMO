using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Maps the documented HTML rich-text subset without interpreting markup as executable content.</summary>
internal static class SpreadsheetXmlRichText {
    internal static readonly XNamespace Html = "http://www.w3.org/TR/REC-html40";
    private static readonly string[] Tags = { "B", "I", "U", "S", "Sub", "Sup", "Font" };

    internal static IReadOnlyList<ExcelRichTextRun> Read(XElement source, SpreadsheetXmlContext context) {
        var result = new List<ExcelRichTextRun>();
        foreach (XNode node in source.Nodes()) Visit(node, new ExcelRichTextRun(string.Empty));
        return result;
        void Visit(XNode node, ExcelRichTextRun inherited) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            if (node is XText text) { result.Add(inherited.Clone(text.Value)); return; }
            if (node is not XElement element) return;
            var run = inherited.Clone();
            if (element.Name.Namespace != Html || !Tags.Contains(element.Name.LocalName)) context.Omitted("rich-text markup " + element.Name);
            else switch (element.Name.LocalName) {
                case "B": run.Bold = true; break;
                case "I": run.Italic = true; break;
                case "U": run.Underline = true; break;
                case "S": run.Strikethrough = true; break;
                case "Sub": run.VerticalTextAlignment = ExcelVerticalTextAlignment.Subscript; break;
                case "Sup": run.VerticalTextAlignment = ExcelVerticalTextAlignment.Superscript; break;
                case "Font":
                    foreach (XAttribute attribute in element.Attributes()) {
                        if (attribute.IsNamespaceDeclaration) continue;
                        if (attribute.Name.Namespace != Html && attribute.Name.Namespace != XNamespace.None) { context.Omitted("rich-text font attribute " + attribute.Name); continue; }
                        switch (attribute.Name.LocalName) {
                            case "Face": run.FontName = attribute.Value; break;
                            case "Size":
                                if (!double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) || double.IsNaN(size) || double.IsInfinity(size) || size <= 0 || size > 409) throw new InvalidDataException("Invalid rich-text font size.");
                                run.FontSize = OfficeLength.Points(size); break;
                            case "Color":
                                string color = attribute.Value;
                                if (color.Length != 7 || color[0] != '#' || !uint.TryParse(color.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) throw new InvalidDataException("Invalid rich-text font color.");
                                run.FontColor = color.Substring(1); break;
                            default: context.Omitted("rich-text font attribute " + attribute.Name); break;
                        }
                    }
                    break;
            }
            if (element.Name.LocalName != "Font" && element.Attributes().Any(item => !item.IsNamespaceDeclaration)) context.Omitted("rich-text tag attributes");
            foreach (XNode child in element.Nodes()) Visit(child, run);
        }
    }

    internal static void Write(XmlWriter writer, OpenXmlCompositeElement source, WorkbookPart workbook, SpreadsheetXmlContext context) {
        foreach (OpenXmlElement item in source.ChildElements) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            if (item is Text text) { writer.WriteString(text.Text); continue; }
            if (item is not Run run) { context.Omitted("phonetic or unsupported rich-text content"); continue; }
            var properties = run.RunProperties;
            var value = ExcelRichTextRun.FromOpenXml(string.Concat(run.Elements<Text>().Select(text => text.Text)), properties,
                Utilities.ExcelThemeColorResolver.Resolve(properties?.GetFirstChild<Color>(), workbook));
            int wrappers = 0;
            if (value.FontName != null || value.FontSize.HasValue || value.FontColor != null) {
                writer.WriteStartElement("html", "Font", Html.NamespaceName); wrappers++;
                if (value.FontName != null) writer.WriteAttributeString("html", "Face", Html.NamespaceName, value.FontName);
                if (value.FontSize.HasValue) writer.WriteAttributeString("html", "Size", Html.NamespaceName, value.FontSize.Value.ToPoints().ToString("R", CultureInfo.InvariantCulture));
                if (value.FontColor is string rgb) writer.WriteAttributeString("html", "Color", Html.NamespaceName, "#" + rgb.Substring(rgb.Length - 6));
            }
            Add(value.Bold, "B"); Add(value.Italic, "I"); Add(value.Underline, "U"); Add(value.Strikethrough, "S");
            if (value.VerticalTextAlignment == ExcelVerticalTextAlignment.Subscript) Add(true, "Sub");
            if (value.VerticalTextAlignment == ExcelVerticalTextAlignment.Superscript) Add(true, "Sup");
            if (value.BoldSpecified && !value.Bold || value.ItalicSpecified && !value.Italic || value.UnderlineSpecified && !value.Underline || value.StrikethroughSpecified && !value.Strikethrough)
                context.Omitted("explicit disabled rich-text emphasis");
            if (value.UnderlineStyle is not null && value.UnderlineStyle != ExcelUnderlineStyle.None && value.UnderlineStyle != ExcelUnderlineStyle.Single) context.Omitted("rich-text underline style (single underline retained)");
            if (value.Outline || value.Shadow || value.Condense || value.Extend) context.Omitted("rich-text font effects");
            writer.WriteString(value.Text);
            while (wrappers-- > 0) writer.WriteEndElement();
            void Add(bool enabled, string name) { if (enabled) { writer.WriteStartElement("html", name, Html.NamespaceName); wrappers++; } }
        }
    }
}
