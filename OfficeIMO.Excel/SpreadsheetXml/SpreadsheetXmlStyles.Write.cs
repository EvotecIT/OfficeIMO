using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static partial class SpreadsheetXmlStyles {
    internal static IReadOnlyList<XElement> Export(WorkbookPart workbook, SpreadsheetXmlContext context) {
        Stylesheet source = workbook.WorkbookStylesPart?.Stylesheet ?? ExcelDocument.CreateDefaultStylesheet();
        foreach (var child in source.ChildElements)
            if (child is not Fonts && child is not Fills && child is not Borders && child is not CellStyleFormats && child is not CellFormats && child is not NumberingFormats && child is not CellStyles && child.HasChildren)
                context.Omitted("stylesheet " + child.LocalName);
        if (source.CellStyles?.Elements<CellStyle>().Any(item => item.BuiltinId?.Value != 0) == true) context.Omitted("named cell-style identity");
        var formats = source.NumberingFormats?.Elements<NumberingFormat>().ToDictionary(item => item.NumberFormatId!.Value, item => item.FormatCode!.Value!) ?? new Dictionary<uint, string>();
        var result = new List<XElement> { new XElement(Ss + "Style", new XAttribute(Ss + "ID", "Default"), WriteFont(source.Fonts!.Elements<Font>().First(), workbook, context)) };
        int styleIndex = 0;
        foreach (CellFormat cell in source.CellFormats!.Elements<CellFormat>()) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            if (result.Count >= context.Options.MaxStyles) throw new InvalidDataException("Workbook exceeds MaxStyles.");
            var style = new XElement(Ss + "Style", new XAttribute(Ss + "ID", StyleId((uint)styleIndex++)));
            uint fontId = cell.FontId?.Value ?? 0;
            var font = source.Fonts!.Elements<Font>().ElementAtOrDefault((int)fontId) ?? throw new InvalidDataException("Invalid cell font index.");
            style.Add(WriteFont(font, workbook, context));
            var fill = source.Fills!.Elements<Fill>().ElementAtOrDefault((int)(cell.FillId?.Value ?? 0)) ?? throw new InvalidDataException("Invalid cell fill index.");
            if (fill.GradientFill != null) context.Omitted("gradient fills");
            if (fill.PatternFill is PatternFill pattern) {
                var interior = new XElement(Ss + "Interior", new XAttribute(Ss + "Pattern", PatternName(pattern.PatternType?.Value ?? PatternValues.None)));
                AddColor(interior, "Color", pattern.ForegroundColor, workbook, context);
                AddColor(interior, "PatternColor", pattern.BackgroundColor, workbook, context); style.Add(interior);
            }
            var border = source.Borders!.Elements<Border>().ElementAtOrDefault((int)(cell.BorderId?.Value ?? 0)) ?? throw new InvalidDataException("Invalid cell border index.");
            style.Add(WriteBorders(border, workbook, context));
            uint number = cell.NumberFormatId?.Value ?? 0;
            string? code = formats.TryGetValue(number, out string? custom) ? custom : ExcelBuiltInNumberFormats.GetCode(number);
            if (code == null) { context.Omitted("unmapped built-in number format " + number); code = "General"; }
            style.Add(new XElement(Ss + "NumberFormat", new XAttribute(Ss + "Format", code)));
            if (cell.Alignment is Alignment alignment) style.Add(WriteAlignment(alignment));
            if (cell.Protection is Protection protection) style.Add(new XElement(Ss + "Protection",
                new XAttribute(Ss + "Protected", protection.Locked?.Value == false ? "0" : "1"), new XAttribute(X + "HideFormula", protection.Hidden?.Value == true ? "1" : "0")));
            if ((cell.FormatId?.Value ?? 0) != 0) context.Omitted("named base-style identity");
            result.Add(style);
        }
        return result;
    }

    private static XElement WriteFont(Font font, WorkbookPart workbook, SpreadsheetXmlContext context) {
        var result = new XElement(Ss + "Font");
        if (font.FontName?.Val?.Value is string name) result.Add(new XAttribute(Ss + "FontName", name));
        if (font.FontSize?.Val?.Value is double size) result.Add(new XAttribute(Ss + "Size", size.ToString("R", CultureInfo.InvariantCulture)));
        result.Add(new XAttribute(Ss + "Bold", ExcelOpenXmlFontProperty.IsEnabled(font.Bold) ? "1" : "0"),
            new XAttribute(Ss + "Italic", ExcelOpenXmlFontProperty.IsEnabled(font.Italic) ? "1" : "0"),
            new XAttribute(Ss + "StrikeThrough", ExcelOpenXmlFontProperty.IsEnabled(font.Strike) ? "1" : "0"));
        if (font.Underline?.Val != null) result.Add(new XAttribute(Ss + "Underline", font.Underline.Val.InnerText switch {
            "single" => "Single", "double" => "Double", "singleAccounting" => "SingleAccounting", "doubleAccounting" => "DoubleAccounting", _ => "None" }));
        else if (font.Underline != null) result.Add(new XAttribute(Ss + "Underline", "Single"));
        if (font.VerticalTextAlignment?.Val != null) result.Add(new XAttribute(Ss + "VerticalAlign", font.VerticalTextAlignment.Val.InnerText switch {
            "superscript" => "Superscript", "subscript" => "Subscript", _ => "None" }));
        AddColor(result, "Color", font.Color, workbook, context);
        string? fontRgb = Utilities.ExcelThemeColorResolver.Resolve(font.Color, workbook);
        if (fontRgb != null && !Enumerable.Range(0, 64).Any(index => Utilities.ExcelThemeColorResolver.Resolve(new Color { Indexed = (uint)index }, null) == fontRgb))
            context.Approximate("Native Excel XML import may quantize font colors outside its legacy indexed palette; the XML retains the requested RGB.");
        if (font.FontFamilyNumbering?.Val != null) result.Add(new XAttribute(X + "Family", font.FontFamilyNumbering.Val.Value switch { 1 => "Roman", 2 => "Swiss", 3 => "Modern", 4 => "Script", 5 => "Decorative", _ => "Automatic" }));
        if (font.FontCharSet?.Val != null) result.Add(new XAttribute(X + "CharSet", font.FontCharSet.Val.Value));
        if (font.Outline != null || font.Shadow != null || font.Condense != null || font.Extend != null) context.Omitted("font effects");
        return result;
    }

    private static XElement WriteAlignment(Alignment value) {
        var result = new XElement(Ss + "Alignment");
        if (value.Horizontal != null) result.Add(new XAttribute(Ss + "Horizontal", value.Horizontal.InnerText switch {
            "left" => "Left", "center" => "Center", "right" => "Right", "fill" => "Fill", "justify" => "Justify",
            "centerContinuous" => "CenterAcrossSelection", "distributed" => "Distributed", _ => "Automatic" }));
        if (value.Vertical != null) result.Add(new XAttribute(Ss + "Vertical", value.Vertical.InnerText switch {
            "top" => "Top", "center" => "Center", "justify" => "Justify", "distributed" => "Distributed", _ => "Bottom" }));
        if (value.WrapText != null) result.Add(new XAttribute(Ss + "WrapText", value.WrapText.Value ? "1" : "0"));
        if (value.ShrinkToFit != null) result.Add(new XAttribute(Ss + "ShrinkToFit", value.ShrinkToFit.Value ? "1" : "0"));
        if (value.Indent != null) result.Add(new XAttribute(Ss + "Indent", value.Indent.Value));
        if (value.TextRotation?.Value == 255) result.Add(new XAttribute(Ss + "VerticalText", "1"));
        else if (value.TextRotation != null) result.Add(new XAttribute(Ss + "Rotate", value.TextRotation.Value > 90 ? 90 - (int)value.TextRotation.Value : (int)value.TextRotation.Value));
        if (value.ReadingOrder != null) result.Add(new XAttribute(Ss + "ReadingOrder", value.ReadingOrder.Value == 1 ? "LeftToRight" : value.ReadingOrder.Value == 2 ? "RightToLeft" : "Context"));
        return result;
    }

    private static void AddColor(XElement target, string name, ColorType? color, WorkbookPart workbook, SpreadsheetXmlContext context) {
        if (color == null || color.Auto?.Value == true) return;
        string? rgb = Utilities.ExcelThemeColorResolver.Resolve(color, workbook);
        if (rgb == null) { context.Omitted("unresolved style color"); return; }
        target.Add(new XAttribute(Ss + name, "#" + rgb.Substring(rgb.Length - 6)));
        if (color.Theme != null || color.Tint != null) context.Omitted("theme color binding (resolved RGB retained)");
    }

    private static string PatternName(PatternValues pattern) => new DocumentFormat.OpenXml.EnumValue<PatternValues>(pattern).InnerText switch {
        "solid" => "Solid", "darkGray" => "Gray75", "mediumGray" => "Gray50", "lightGray" => "Gray25", "gray125" => "Gray125", "gray0625" => "Gray0625",
        "darkHorizontal" => "HorzStripe", "darkVertical" => "VertStripe", "darkDown" => "ReverseDiagStripe", "darkUp" => "DiagStripe",
        "darkGrid" => "DiagCross", "darkTrellis" => "ThickDiagCross", "lightHorizontal" => "ThinHorzStripe", "lightVertical" => "ThinVertStripe",
        "lightDown" => "ThinReverseDiagStripe", "lightUp" => "ThinDiagStripe", "lightGrid" => "ThinHorzCross", "lightTrellis" => "ThinDiagCross", _ => "None" };
}
