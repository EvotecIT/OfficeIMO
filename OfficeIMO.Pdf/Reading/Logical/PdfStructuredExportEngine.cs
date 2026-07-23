using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Pdf;

/// <summary>Exports one logical PDF model to JSON, Markdown, ALTO XML, hOCR, or PAGE XML.</summary>
internal static class PdfStructuredExportEngine {
    private const string JsonSchema = "officeimo.pdf.logical.v1";
    private const string AltoNamespace = "http://www.loc.gov/standards/alto/ns-v4#";
    private const string PageNamespace = "http://schema.primaresearch.org/PAGE/gts/pagecontent/2019-07-15";

    /// <summary>Exports an already parsed logical document without rerunning extraction.</summary>
    public static string Export(PdfLogicalDocument document, PdfStructuredExportFormat format) {
        Guard.NotNull(document, nameof(document));
        switch (format) {
            case PdfStructuredExportFormat.Json:
                return ToJson(document);
            case PdfStructuredExportFormat.Markdown:
                return document.ToMarkdown();
            case PdfStructuredExportFormat.AltoXml:
                return ToAltoXml(document);
            case PdfStructuredExportFormat.Hocr:
                return ToHocr(document);
            case PdfStructuredExportFormat.PageXml:
                return ToPageXml(document);
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported PDF structured export format.");
        }
    }

    /// <summary>Loads PDF bytes once into the logical model and exports the requested format.</summary>
    public static string Export(byte[] pdf, PdfStructuredExportFormat format, PdfTextLayoutOptions? layoutOptions = null) =>
        Export(pdf, format, layoutOptions, readOptions: null);

    /// <summary>Loads PDF bytes with explicit read limits or credentials and exports the requested format.</summary>
    public static string Export(byte[] pdf, PdfStructuredExportFormat format, PdfTextLayoutOptions? layoutOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return Export(PdfLogicalDocument.Load(pdf, layoutOptions, readOptions), format);
    }

    /// <summary>
    /// Exports one schema-valid PAGE XML document per logical page. PAGE XML is image/page scoped and does not define a multi-page root.
    /// </summary>
    public static IReadOnlyList<string> ExportPageXmlDocuments(PdfLogicalDocument document) {
        Guard.NotNull(document, nameof(document));
        var results = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            results.Add(BuildPageXml(document.Pages[i], i + 1));
        }

        return results.AsReadOnly();
    }

    private static string ToJson(PdfLogicalDocument document) {
        var builder = new StringBuilder();
        builder.Append("{\"schema\":\"").Append(JsonSchema).Append("\",\"pages\":[");
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            if (pageIndex > 0) builder.Append(',');
            PdfLogicalPage page = document.Pages[pageIndex];
            builder.Append("{\"number\":").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture))
                .Append(",\"width\":").Append(Number(page.Width))
                .Append(",\"height\":").Append(Number(page.Height))
                .Append(",\"rotation\":").Append(page.RotationDegrees.ToString(CultureInfo.InvariantCulture))
                .Append(",\"lines\":[");
            IReadOnlyList<ExportLine> lines = GetLines(page);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                if (lineIndex > 0) builder.Append(',');
                ExportLine line = lines[lineIndex];
                builder.Append("{\"order\":").Append((lineIndex + 1).ToString(CultureInfo.InvariantCulture))
                    .Append(",\"text\":\"").Append(Json(line.Text)).Append('"')
                    .Append(",\"x\":").Append(Number(line.X))
                    .Append(",\"y\":").Append(Number(line.Y))
                    .Append(",\"width\":").Append(Number(line.Width))
                    .Append(",\"height\":").Append(Number(line.Height))
                    .Append('}');
            }

            builder.Append("]}");
        }

        builder.Append("]}");
        return builder.ToString();
    }

    private static string ToAltoXml(PdfLogicalDocument document) {
        XNamespace alto = AltoNamespace;
        var layout = new XElement(alto + "Layout");
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfLogicalPage page = document.Pages[pageIndex];
            var textBlock = new XElement(alto + "TextBlock", new XAttribute("ID", "block_" + (pageIndex + 1)));
            IReadOnlyList<ExportLine> lines = GetLines(page);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                ExportLine line = lines[lineIndex];
                textBlock.Add(new XElement(
                    alto + "TextLine",
                    new XAttribute("ID", "line_" + (pageIndex + 1) + "_" + (lineIndex + 1)),
                    BoxAttributes(line),
                    new XElement(
                        alto + "String",
                        new XAttribute("ID", "string_" + (pageIndex + 1) + "_" + (lineIndex + 1)),
                        new XAttribute("CONTENT", XmlText(line.Text)),
                        BoxAttributes(line))));
            }

            layout.Add(new XElement(
                alto + "Page",
                new XAttribute("ID", "page_" + (pageIndex + 1)),
                new XAttribute("PHYSICAL_IMG_NR", page.PageNumber),
                new XAttribute("WIDTH", Number(page.Width)),
                new XAttribute("HEIGHT", Number(page.Height)),
                new XElement(alto + "PrintSpace", textBlock)));
        }

        var documentElement = new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(
                alto + "alto",
                new XAttribute("SCHEMAVERSION", "4.4"),
                new XElement(alto + "Description", new XElement(alto + "MeasurementUnit", "point")),
                layout));
        return documentElement.ToString(SaveOptions.DisableFormatting);
    }

    private static string ToHocr(PdfLogicalDocument document) {
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        var body = new XElement(xhtml + "body");
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfLogicalPage page = document.Pages[pageIndex];
            var pageElement = new XElement(
                xhtml + "div",
                new XAttribute("class", "ocr_page"),
                new XAttribute("id", "page_" + (pageIndex + 1)),
                new XAttribute("title", "bbox 0 0 " + Integer(page.Width) + " " + Integer(page.Height)));
            IReadOnlyList<ExportLine> lines = GetLines(page);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                ExportLine line = lines[lineIndex];
                pageElement.Add(new XElement(
                    xhtml + "span",
                    new XAttribute("class", "ocr_line"),
                    new XAttribute("id", "line_" + (pageIndex + 1) + "_" + (lineIndex + 1)),
                    new XAttribute("title", "bbox " + Integer(line.X) + " " + Integer(line.Y) + " " + Integer(line.X + line.Width) + " " + Integer(line.Y + line.Height)),
                    XmlText(line.Text)));
            }

            body.Add(pageElement);
        }

        return new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(
                xhtml + "html",
                new XElement(xhtml + "head", new XElement(xhtml + "meta", new XAttribute("name", "ocr-system"), new XAttribute("content", "OfficeIMO.Pdf"))),
                body)).ToString(SaveOptions.DisableFormatting);
    }

    private static string ToPageXml(PdfLogicalDocument document) {
        if (document.Pages.Count != 1) {
            throw new InvalidOperationException(
                "PAGE XML is page scoped. Use ToPageXmlDocuments to export one schema-valid document per logical page.");
        }

        return BuildPageXml(document.Pages[0], 1);
    }

    private static string BuildPageXml(PdfLogicalPage page, int pageIndex) {
        XNamespace pageNs = PageNamespace;
        var region = new XElement(pageNs + "TextRegion", new XAttribute("id", "region_" + pageIndex));
        IReadOnlyList<ExportLine> lines = GetLines(page);
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            ExportLine line = lines[lineIndex];
            region.Add(new XElement(
                pageNs + "TextLine",
                new XAttribute("id", "line_" + pageIndex + "_" + (lineIndex + 1)),
                new XElement(pageNs + "Coords", new XAttribute("points", Points(line))),
                new XElement(pageNs + "TextEquiv", new XElement(pageNs + "Unicode", XmlText(line.Text)))));
        }

        return new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(
                pageNs + "PcGts",
                new XElement(
                    pageNs + "Metadata",
                    new XElement(pageNs + "Creator", "OfficeIMO.Pdf")),
                new XElement(
                    pageNs + "Page",
                    new XAttribute("imageFilename", "page-" + page.PageNumber + ".png"),
                    new XAttribute("imageWidth", Integer(page.Width)),
                    new XAttribute("imageHeight", Integer(page.Height)),
                    region))).ToString(SaveOptions.DisableFormatting);
    }

    private static List<ExportLine> GetLines(PdfLogicalPage page) {
        var lines = new List<ExportLine>(page.TextBlocks.Count);
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfLogicalTextBlock block = page.TextBlocks[i];
            double height = Math.Max(1D, block.FontSize * 1.2D);
            double y = Math.Max(0D, page.Height - block.BaselineY - block.FontSize);
            lines.Add(new ExportLine(
                block.Text,
                Math.Max(0D, block.XStart),
                y,
                Math.Max(0D, block.XEnd - block.XStart),
                Math.Min(height, Math.Max(0D, page.Height - y))));
        }

        return lines;
    }

    private static IEnumerable<XAttribute> BoxAttributes(ExportLine line) {
        yield return new XAttribute("HPOS", Number(line.X));
        yield return new XAttribute("VPOS", Number(line.Y));
        yield return new XAttribute("WIDTH", Number(line.Width));
        yield return new XAttribute("HEIGHT", Number(line.Height));
    }

    private static string Points(ExportLine line) {
        return Integer(line.X) + "," + Integer(line.Y) + " " +
            Integer(line.X + line.Width) + "," + Integer(line.Y) + " " +
            Integer(line.X + line.Width) + "," + Integer(line.Y + line.Height) + " " +
            Integer(line.X) + "," + Integer(line.Y + line.Height);
    }

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string Integer(double value) => Math.Max(0, (int)Math.Ceiling(value)).ToString(CultureInfo.InvariantCulture);

    private static string XmlText(string value) {
        StringBuilder? builder = null;
        for (int i = 0; i < value.Length; i++) {
            char current = value[i];
            bool valid = current == '\t' || current == '\n' || current == '\r' ||
                current >= ' ' && current <= '\uD7FF' ||
                current >= '\uE000' && current <= '\uFFFD';
            if (valid) {
                builder?.Append(current);
                continue;
            }
            if (char.IsHighSurrogate(current) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) {
                if (builder != null) {
                    builder.Append(current).Append(value[++i]);
                } else {
                    i++;
                }
                continue;
            }
            if (builder == null) {
                builder = new StringBuilder(value.Length);
                builder.Append(value, 0, i);
            }
            builder.Append('\uFFFD');
        }
        return builder?.ToString() ?? value;
    }

    private static string Json(string value) {
        var builder = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\b': builder.Append("\\b"); break;
                case '\f': builder.Append("\\f"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (ch < 32) {
                        builder.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        builder.Append(ch);
                    }
                    break;
            }
        }

        return builder.ToString();
    }

    private readonly struct ExportLine {
        internal ExportLine(string text, double x, double y, double width, double height) {
            Text = text;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        internal string Text { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
    }
}
