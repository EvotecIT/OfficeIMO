using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private static Dictionary<int, string> ReadTextBackgroundPalette(IEnumerable<XElement> entries) {
        var colors = new Dictionary<int, string>();
        foreach (XElement entry in entries) {
            if (entry.Name.LocalName == "ColorEntry" &&
                int.TryParse((string?)entry.Attribute("IX"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
                (string?)entry.Attribute("RGB") is string rgb && !colors.ContainsKey(index)) colors.Add(index, rgb);
        }
        return colors;
    }

    private static Color ParseTextBackgroundColor(string? value, IReadOnlyDictionary<int, string>? colors) {
        string? normalized = NormalizeCellLiteral(value);
        if (TryParseCellIntValue(normalized, out int index)) {
            // TextBkgnd reserves zero (and the legacy sentinel 255) for no fill;
            // its other palette values are one greater than ordinary color cells.
            if (index == 0 || index == 255) return Color.Transparent;
            string paletteIndex = (index - 1).ToString(CultureInfo.InvariantCulture);
            return ParseColor(colors != null && colors.TryGetValue(index - 1, out string? rgb) ? rgb : paletteIndex, default);
        }
        if (normalized?.StartsWith("RGB(", StringComparison.OrdinalIgnoreCase) == true && normalized.EndsWith(")+1", StringComparison.Ordinal))
            normalized = normalized.Substring(0, normalized.Length - 2);
        return ParseColor(normalized, default);
    }

    private static void LoadTextBackgroundColor(VisioTextStyle style, XElement cell, IReadOnlyDictionary<int, string>? colors) {
        style.BackgroundColor = ParseTextBackgroundColor((string?)cell.Attribute("V"), colors);
        style.NativeBackgroundColorCell = new XElement(cell);
        style.BackgroundColorAssigned = false;
    }

    private static void LoadTextBackgroundTransparency(VisioTextStyle style, XElement cell) {
        // Native percentage caches are normalized: 0.25 means 25%, regardless of U/F.
        style.BackgroundTransparency = ParseDouble((string?)cell.Attribute("V")) * 100D;
        style.NativeBackgroundTransparencyCell = new XElement(cell);
        style.BackgroundTransparencyAssigned = false;
    }

    private void WriteTextBackgroundColorCell(XmlWriter writer, string ns, VisioTextStyle? style) {
        if (style?.BackgroundColor is not Color color) return;
        XElement? source = style.NativeBackgroundColorCell;
        // Copies retain native syntax only while its palette index still resolves to
        // the model's color in the destination document. No-fill sentinels are portable.
        if (source != null && ParseTextBackgroundColor((string?)source.Attribute("V"), ReadTextBackgroundPalette(PreservedColorsElements)) == color) {
            source.WriteTo(writer);
            return;
        }
        WriteCanonicalTextBackgroundColor(writer, ns, color);
    }

    private static void WriteCanonicalTextBackgroundColor(XmlWriter writer, string ns, Color color) {
        if (color.A == 0) WriteCell(writer, ns, "TextBkgnd", 0);
        else WriteCellValue(writer, ns, "TextBkgnd", color.ToVisioHex(), null, color.ToVisioRgb() + "+1");
    }

    private static void WriteTextBackgroundTransparencyCell(XmlWriter writer, string ns, VisioTextStyle? style) {
        if (style?.BackgroundTransparency is not double transparency) return;
        if (style.NativeBackgroundTransparencyCell is XElement source) source.WriteTo(writer);
        else WriteCell(writer, ns, "TextBkgndTrans", transparency / 100D);
    }

    private void NormalizeImportedTextBackgroundColors(XDocument source, IReadOnlyDictionary<int, string> sourceColors) {
        var destinationColors = ReadTextBackgroundPalette(PreservedColorsElements);
        foreach (XElement cell in source.Descendants(XName.Get("Cell", VisioNamespace)).Where(cell => (string?)cell.Attribute("N") == "TextBkgnd")) {
            if (!TryParseCellIntValue((string?)cell.Attribute("V"), out int index) || index <= 0 || index == 255) continue;
            Color color = ParseTextBackgroundColor((string?)cell.Attribute("V"), sourceColors);
            if (color == ParseTextBackgroundColor((string?)cell.Attribute("V"), destinationColors)) continue;
            // A conflicting destination palette cannot carry the source index or its formula.
            cell.SetAttributeValue("V", color.ToVisioHex());
            cell.SetAttributeValue("F", color.ToVisioRgb() + "+1");
            cell.Attribute("E")?.Remove();
        }
    }
}
