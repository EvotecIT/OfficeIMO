using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsFontTable {
    // Generated binary streaming styles have one body font with the workbook text color.
    internal static LegacyXlsFontTable CreateStreaming(OfficeXmlNode font, OfficeTheme theme, out ushort fontIndex) {
        const int main = OfficeXmlNamespaces.SpreadsheetML;
        var table = new LegacyXlsFontTable();
        string? authoredName = font.Child(main, "name").Attribute("val");
        string name = string.IsNullOrWhiteSpace(authoredName) ? DefaultFontName : authoredName!;
        if (name.Length > byte.MaxValue) throw new NotSupportedException("Native XLS saving does not yet support cell font names longer than 255 characters.");

        double size = font.Child(main, "sz").DoubleAttribute("val") ?? DefaultFontSize;
        if (size <= 0 || size * 20d > ushort.MaxValue) throw new NotSupportedException("Native XLS saving does not yet support cell font sizes outside BIFF8 limits.");

        ushort colorIndex = AutomaticColorIndex;
        if (font.Child(main, "color").Exists
            && !table.TryMapArgbColor("FF" + theme.Palette.Text.Resolve(theme).ToRgbHex(), out colorIndex, out string? reason)) {
            throw new NotSupportedException($"Native XLS saving does not yet support {reason ?? "cell font colors"}.");
        }

        byte family = checked((byte)(font.Child(main, "family").UInt32Attribute("val") ?? 0));
        fontIndex = table.GetOrAddFont(new FontKey(name.Trim(), ToTwips(size), colorIndex, bold: false, italic: false,
            underlineStyle: 0, strikeout: false, outline: false, shadow: false, condense: false, extend: false,
            escapement: LegacyXlsFontEscapement.None, family: family, characterSet: 1));
        return table;
    }
}
