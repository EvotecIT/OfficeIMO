using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static partial class SpreadsheetXmlStyles {
    private static void ReadAppearance(XElement source, CellFormat cell, Stylesheet styles, SpreadsheetXmlContext context) {
        if (source.Element(Ss + "Font") is XElement fontXml) {
            context.Inspect(fontXml, Array.Empty<string>(), new[] { "FontName", "Size", "Color", "Bold", "Italic", "Underline", "StrikeThrough", "VerticalAlign", (X + "CharSet").ToString(), (X + "Family").ToString() });
            var font = (Font)styles.Fonts!.Elements<Font>().ElementAt((int)(cell.FontId?.Value ?? 0)).CloneNode(true);
            if (Text(fontXml, "FontName") is string name) font.FontName = new FontName { Val = name };
            if (Text(fontXml, "Size") is string size) font.FontSize = new FontSize { Val = Positive(size, 409) };
            if (Text(fontXml, "Color") is string color) font.Color = new Color { Rgb = Color(color) };
            if (Text(fontXml, "Bold") is string bold) font.Bold = new Bold { Val = Flag(bold) };
            if (Text(fontXml, "Italic") is string italic) font.Italic = new Italic { Val = Flag(italic) };
            if (Text(fontXml, "StrikeThrough") is string strike) font.Strike = new Strike { Val = Flag(strike) };
            if (Text(fontXml, "Underline") is string underline) font.Underline = new Underline { Val = underline switch {
                "None" => UnderlineValues.None, "Single" => UnderlineValues.Single, "Double" => UnderlineValues.Double,
                "SingleAccounting" => UnderlineValues.SingleAccounting, "DoubleAccounting" => UnderlineValues.DoubleAccounting,
                _ => throw new InvalidDataException("Invalid font underline.") } };
            if (Text(fontXml, "VerticalAlign") is string vertical) font.VerticalTextAlignment = new VerticalTextAlignment { Val = vertical switch {
                "Subscript" => VerticalAlignmentRunValues.Subscript, "Superscript" => VerticalAlignmentRunValues.Superscript,
                "None" => VerticalAlignmentRunValues.Baseline, _ => throw new InvalidDataException("Invalid font vertical alignment.") } };
            if ((string?)fontXml.Attribute(X + "CharSet") is string charset) font.FontCharSet = new FontCharSet { Val = int.Parse(charset, CultureInfo.InvariantCulture) };
            if ((string?)fontXml.Attribute(X + "Family") is string family) font.FontFamilyNumbering = new FontFamilyNumbering { Val = family switch { "Roman" => 1, "Swiss" => 2, "Modern" => 3, "Script" => 4, "Decorative" => 5, "Automatic" => 0, _ => throw new InvalidDataException("Invalid font family classification.") } };
            cell.FontId = AppendUnique(styles.Fonts, font); cell.ApplyFont = true;
        }
        if (source.Element(Ss + "Interior") is XElement interior) {
            context.Inspect(interior, Array.Empty<string>(), new[] { "Color", "Pattern", "PatternColor" });
            var fill = (Fill)styles.Fills!.Elements<Fill>().ElementAt((int)(cell.FillId?.Value ?? 0)).CloneNode(true);
            fill.PatternFill ??= new PatternFill { PatternType = PatternValues.None };
            string? pattern = Text(interior, "Pattern");
            if (pattern != null) fill.PatternFill.PatternType = Pattern(pattern);
            if (Text(interior, "Color") is string color) fill.PatternFill.ForegroundColor = new ForegroundColor { Rgb = Color(color) };
            if (Text(interior, "PatternColor") is string background) fill.PatternFill.BackgroundColor = new BackgroundColor { Rgb = Color(background) };
            cell.FillId = AppendUnique(styles.Fills, fill); cell.ApplyFill = true;
        }
        if (source.Element(Ss + "Alignment") is XElement alignment) {
            context.Inspect(alignment, Array.Empty<string>(), new[] { "Horizontal", "Vertical", "WrapText", "ShrinkToFit", "Indent", "Rotate", "VerticalText", "ReadingOrder" });
            var value = cell.Alignment ?? new Alignment();
            if (Text(alignment, "Horizontal") is string horizontal) value.Horizontal = horizontal switch {
                "Automatic" => HorizontalAlignmentValues.General, "Left" => HorizontalAlignmentValues.Left, "Center" => HorizontalAlignmentValues.Center,
                "Right" => HorizontalAlignmentValues.Right, "Fill" => HorizontalAlignmentValues.Fill, "Justify" => HorizontalAlignmentValues.Justify,
                "CenterAcrossSelection" => HorizontalAlignmentValues.CenterContinuous, "Distributed" => HorizontalAlignmentValues.Distributed,
                _ => throw new InvalidDataException("Invalid horizontal alignment.") };
            if (Text(alignment, "Vertical") is string vertical) value.Vertical = vertical switch {
                "Top" => VerticalAlignmentValues.Top, "Center" => VerticalAlignmentValues.Center, "Bottom" => VerticalAlignmentValues.Bottom,
                "Justify" => VerticalAlignmentValues.Justify, "Distributed" => VerticalAlignmentValues.Distributed,
                _ => throw new InvalidDataException("Invalid vertical alignment.") };
            if (Text(alignment, "WrapText") is string wrap) value.WrapText = Flag(wrap);
            if (Text(alignment, "ShrinkToFit") is string shrink) value.ShrinkToFit = Flag(shrink);
            if (Text(alignment, "Indent") is string indent) value.Indent = (uint)Nonnegative(indent, 250);
            if (Text(alignment, "Rotate") is string rotate) {
                if (!int.TryParse(rotate, NumberStyles.Integer, CultureInfo.InvariantCulture, out int degrees) || degrees < -90 || degrees > 90) throw new InvalidDataException("Invalid text rotation.");
                value.TextRotation = (uint)(degrees < 0 ? 90 - degrees : degrees);
            }
            if (Text(alignment, "VerticalText") is string stacked && Flag(stacked)) value.TextRotation = 255;
            if (Text(alignment, "ReadingOrder") is string order) value.ReadingOrder = order switch {
                "Context" => 0U, "LeftToRight" => 1U, "RightToLeft" => 2U, _ => throw new InvalidDataException("Invalid reading order.") };
            cell.Alignment = value; cell.ApplyAlignment = true;
        }
        if (source.Element(Ss + "Protection") is XElement protection) {
            context.Inspect(protection, Array.Empty<string>(), new[] { "Protected", (X + "HideFormula").ToString() });
            cell.Protection ??= new Protection();
            if (Text(protection, "Protected") is string locked) cell.Protection.Locked = Flag(locked);
            if ((string?)protection.Attribute(X + "HideFormula") is string hidden) cell.Protection.Hidden = Flag(hidden);
            cell.ApplyProtection = true;
        }
        if (source.Element(Ss + "Borders") is XElement borders) ReadBorders(borders, cell, styles, context);
    }

    private static string? Text(XElement element, string name) => (string?)element.Attribute(Ss + name);
    private static bool Flag(string text) => text switch { "1" or "true" => true, "0" or "false" => false, _ => throw new InvalidDataException("Invalid XML Spreadsheet Boolean attribute.") };
    private static double Positive(string text, double maximum) => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
        && !double.IsInfinity(value) && value > 0 && value <= maximum ? value : throw new InvalidDataException("Invalid XML Spreadsheet dimension.");
    private static int Nonnegative(string text, int maximum) => int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value)
        && value >= 0 && value <= maximum ? value : throw new InvalidDataException("Invalid XML Spreadsheet integer attribute.");
    private static string Color(string text) => text.Length == 7 && text[0] == '#' && uint.TryParse(text.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)
        ? "FF" + text.Substring(1).ToUpperInvariant() : throw new InvalidDataException("Invalid XML Spreadsheet RGB color.");
    private static PatternValues Pattern(string name) => name switch {
        "None" => PatternValues.None, "Solid" => PatternValues.Solid, "Gray75" => PatternValues.DarkGray, "Gray50" => PatternValues.MediumGray,
        "Gray25" => PatternValues.LightGray, "Gray125" => PatternValues.Gray125, "Gray0625" => PatternValues.Gray0625,
        "HorzStripe" => PatternValues.DarkHorizontal, "VertStripe" => PatternValues.DarkVertical, "ReverseDiagStripe" => PatternValues.DarkDown,
        "DiagStripe" => PatternValues.DarkUp, "DiagCross" => PatternValues.DarkGrid, "ThickDiagCross" => PatternValues.DarkTrellis,
        "ThinHorzStripe" => PatternValues.LightHorizontal, "ThinVertStripe" => PatternValues.LightVertical, "ThinReverseDiagStripe" => PatternValues.LightDown,
        "ThinDiagStripe" => PatternValues.LightUp, "ThinHorzCross" => PatternValues.LightGrid, "ThinDiagCross" => PatternValues.LightTrellis,
        _ => throw new InvalidDataException("Invalid XML Spreadsheet fill pattern.") };
}
