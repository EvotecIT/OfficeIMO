using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Converts XML point dimensions through the shared workbook-font geometry owner.</summary>
internal static class SpreadsheetXmlGeometry {
    internal static uint?[] ReadColumns(XElement table, ExcelSheet sheet, IReadOnlyDictionary<string, uint> styles, SpreadsheetXmlContext context) {
        var result = new uint?[257];
        var worksheet = sheet.WorksheetPart.Worksheet!;
        var format = worksheet.GetFirstChild<SheetFormatProperties>() ?? new SheetFormatProperties { DefaultRowHeight = 15 };
        double digit = table.Attribute(Ss + "DefaultColumnWidth") != null || table.Elements(Ss + "Column").Any(item => item.Attribute(Ss + "Width") != null) ? DigitWidth(sheet.WorksheetPart, context) : 7;
        if (table.Attribute(Ss + "DefaultColumnWidth") is XAttribute width) format.DefaultColumnWidth = Width(width.Value, digit);
        if (table.Attribute(Ss + "DefaultRowHeight") is XAttribute height) format.DefaultRowHeight = Dimension(height.Value, 409);
        worksheet.AddChild(format, true);
        var columns = new Columns(); int previous = 0;
        foreach (XElement source in table.Elements(Ss + "Column")) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            context.Inspect(source, Array.Empty<string>(), new[] { "Index", "Span", "Width", "AutoFitWidth", "Hidden", "StyleID" });
            int start = Index(source, "Index", previous + 1, 256);
            int span = (string?)source.Attribute(Ss + "Span") is string spanText && int.TryParse(spanText, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ? parsed : source.Attribute(Ss + "Span") == null ? 0 : -1;
            if (span < 0 || span > 255 || start <= previous) throw new InvalidDataException("Invalid XML Spreadsheet column span or order.");
            int end = start + span; Coordinate(1, end); previous = end;
            var column = new Column { Min = (uint)start, Max = (uint)end };
            if (source.Attribute(Ss + "Width") is XAttribute size) { column.Width = Width(size.Value, digit); column.CustomWidth = true; }
            if (source.Attribute(Ss + "Hidden") is XAttribute hidden) column.Hidden = Flag(hidden.Value);
            if (source.Attribute(Ss + "StyleID") is XAttribute style) {
                if (!styles.TryGetValue(style.Value, out uint id)) throw new InvalidDataException("Unknown column style.");
                column.Style = id; for (int index = start; index <= end; index++) result[index] = id;
            }
            columns.Append(column);
        }
        if (columns.HasChildren) worksheet.AddChild(columns, true);
        return result;
    }

    internal static void WriteColumns(XmlWriter writer, WorksheetPart part, int styleCount, SpreadsheetXmlContext context) {
        var format = part.Worksheet!.GetFirstChild<SheetFormatProperties>();
        double digit = format != null || part.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().Any(item => item.Width != null) == true ? DigitWidth(part, context) : 7;
        if (format != null) {
            if (format.DefaultRowHeight != null) Attribute(writer, "DefaultRowHeight", format.DefaultRowHeight.Value);
            Attribute(writer, "DefaultColumnWidth", ExcelColumnGeometry.GetWidthPixels(null, format, digit, true) * 72d / 96d);
        }
        uint previous = 0;
        foreach (Column column in part.Worksheet.GetFirstChild<Columns>()?.Elements<Column>() ?? Enumerable.Empty<Column>()) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            uint start = column.Min?.Value ?? 0, end = column.Max?.Value ?? 0;
            if (start < 1 || end < start || end > A1.MaxColumns || start <= previous) throw new InvalidDataException("XML Spreadsheet requires ordered nonoverlapping columns within 256.");
            previous = end;
            if (end > 256) { context.Omitted("column formatting beyond XML Spreadsheet column 256"); if (start > 256) continue; end = 256; }
            writer.WriteStartElement("Column", Ss.NamespaceName); Attribute(writer, "Index", start);
            if (end > start) Attribute(writer, "Span", end - start);
            if (column.Width != null) Attribute(writer, "Width", ExcelColumnGeometry.GetWidthPixels(column, format, digit, true) * 72d / 96d);
            if (column.Hidden?.Value == true) Attribute(writer, "Hidden", 1);
            if (column.Style != null) {
                if (column.Style.Value >= styleCount) throw new InvalidDataException("Invalid column style.");
                writer.WriteAttributeString("ss", "StyleID", Ss.NamespaceName, SpreadsheetXmlStyles.StyleId(column.Style.Value));
            }
            if (column.OutlineLevel?.Value > 0 || column.Collapsed?.Value == true) context.Omitted("column outlining");
            writer.WriteEndElement();
        }
    }

    private static double DigitWidth(WorksheetPart part, SpreadsheetXmlContext context) {
        if (context.Options.DefaultFontMaximumDigitWidthPixels is double measured) return measured;
        double digit = ExcelColumnGeometry.GetMaximumDigitWidth(part, out bool resolved);
        if (!resolved) context.Approximate("Default font metrics are unavailable; column point conversion uses a seven-pixel digit width at 96 DPI.");
        return digit;
    }
    internal static double Dimension(string text, double maximum) => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
        && !double.IsInfinity(value) && !double.IsNaN(value) && value >= 0 && value <= maximum ? value : throw new InvalidDataException("Invalid XML Spreadsheet dimension.");
    private static double Width(string text, double digit) {
        double points = Dimension(text, 65535);
        double width = Math.Round(points * 96d / 72d / digit * 256d) / 256d;
        return width <= 255 ? width : throw new InvalidDataException("Column width exceeds the Open XML limit for this workbook font.");
    }
    internal static bool Flag(string text) => text switch { "1" or "true" or "True" => true, "0" or "false" or "False" => false, _ => throw new InvalidDataException("Invalid XML Spreadsheet Boolean.") };
    private static void Attribute(XmlWriter writer, string name, double value) => writer.WriteAttributeString("ss", name, Ss.NamespaceName, value.ToString("R", CultureInfo.InvariantCulture));
}
