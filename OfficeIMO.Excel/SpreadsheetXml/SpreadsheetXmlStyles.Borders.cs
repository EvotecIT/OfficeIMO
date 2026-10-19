using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static partial class SpreadsheetXmlStyles {
    private static void ReadBorders(XElement source, CellFormat cell, Stylesheet styles, SpreadsheetXmlContext context) {
        context.Inspect(source, new[] { "Border" }, Array.Empty<string>());
        var border = (Border)styles.Borders!.Elements<Border>().ElementAt((int)(cell.BorderId?.Value ?? 0)).CloneNode(true);
        foreach (XElement item in source.Elements(Ss + "Border")) {
            context.Inspect(item, Array.Empty<string>(), new[] { "Position", "LineStyle", "Weight", "Color" });
            string position = Text(item, "Position") ?? throw new InvalidDataException("Missing border position.");
            BorderPropertiesType edge = position switch {
                "Left" => border.LeftBorder ??= new LeftBorder(), "Right" => border.RightBorder ??= new RightBorder(),
                "Top" => border.TopBorder ??= new TopBorder(), "Bottom" => border.BottomBorder ??= new BottomBorder(),
                "DiagonalLeft" or "DiagonalRight" => border.DiagonalBorder ??= new DiagonalBorder(),
                _ => throw new InvalidDataException("Invalid border position.") };
            if (position == "DiagonalLeft") border.DiagonalDown = true;
            if (position == "DiagonalRight") border.DiagonalUp = true;
            int weight = Nonnegative(Text(item, "Weight") ?? "1", 3);
            if (Text(item, "LineStyle") is string line) edge.Style = BorderStyle(line, weight);
            if (Text(item, "Color") is string color) edge.Color = new Color { Rgb = Color(color) };
        }
        cell.BorderId = AppendUnique(styles.Borders, border); cell.ApplyBorder = true;
    }

    private static BorderStyleValues BorderStyle(string line, int weight) => line switch {
        "None" => BorderStyleValues.None,
        "Continuous" => weight == 0 ? BorderStyleValues.Hair : weight == 1 ? BorderStyleValues.Thin : weight == 2 ? BorderStyleValues.Medium : BorderStyleValues.Thick,
        "Dash" => weight >= 2 ? BorderStyleValues.MediumDashed : BorderStyleValues.Dashed,
        "Dot" => BorderStyleValues.Dotted,
        "DashDot" => weight >= 2 ? BorderStyleValues.MediumDashDot : BorderStyleValues.DashDot,
        "DashDotDot" => weight >= 2 ? BorderStyleValues.MediumDashDotDot : BorderStyleValues.DashDotDot,
        "Double" => BorderStyleValues.Double, "SlantDashDot" => BorderStyleValues.SlantDashDot,
        _ => throw new InvalidDataException("Invalid XML Spreadsheet border line style.") };

    private static XElement WriteBorders(Border border, DocumentFormat.OpenXml.Packaging.WorkbookPart workbook, SpreadsheetXmlContext context) {
        var result = new XElement(Ss + "Borders");
        Add(border.LeftBorder, "Left"); Add(border.RightBorder, "Right"); Add(border.TopBorder, "Top"); Add(border.BottomBorder, "Bottom");
        if (border.DiagonalDown?.Value == true) Add(border.DiagonalBorder, "DiagonalLeft");
        if (border.DiagonalUp?.Value == true) Add(border.DiagonalBorder, "DiagonalRight");
        if (border.StartBorder != null || border.EndBorder != null || border.VerticalBorder != null || border.HorizontalBorder != null)
            context.Omitted("logical or internal borders");
        return result;

        void Add(BorderPropertiesType? edge, string position) {
            if (edge?.Style == null || edge.Style.Value == BorderStyleValues.None) return;
            string style = edge.Style.InnerText!;
            string line = style switch {
                "dashed" or "mediumDashed" => "Dash", "dotted" => "Dot", "dashDot" or "mediumDashDot" => "DashDot",
                "dashDotDot" or "mediumDashDotDot" => "DashDotDot", "double" => "Double", "slantDashDot" => "SlantDashDot", _ => "Continuous" };
            int weight = style == "hair" ? 0 : style == "thick" ? 3 : style.StartsWith("medium", StringComparison.Ordinal) ? 2 : 1;
            var item = new XElement(Ss + "Border", new XAttribute(Ss + "Position", position), new XAttribute(Ss + "LineStyle", line),
                new XAttribute(Ss + "Weight", weight.ToString(CultureInfo.InvariantCulture)));
            AddColor(item, "Color", edge.Color, workbook, context); result.Add(item);
        }
    }
}
