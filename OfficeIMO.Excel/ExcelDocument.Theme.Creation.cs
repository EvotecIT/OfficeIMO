using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private OfficeTheme? _creationTheme;
    /// <summary>Gets the creation palette for status colors that have no native theme slot.</summary>
    internal OfficeTheme EffectiveCreationTheme => _creationTheme ?? OfficeTheme.Default;

    private static string CreateWorkbookThemeXml(OfficeTheme theme) {
        var xml = XDocument.Parse(Encoding.UTF8.GetString(DefaultThemeBytes.Value));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var root = xml.Root!;
        root.SetAttributeValue("name", theme.Name);
        var elements = root.Element(a + "themeElements")!;
        var colors = elements.Element(a + "clrScheme")!;
        colors.SetAttributeValue("name", theme.Name);
        var palette = theme.Palette;
        var slots = new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        var values = new[] { palette.Text, palette.Background, palette.TextMuted, palette.Background,
            palette.Accent1, palette.Accent2, palette.Accent3, palette.Accent4, palette.Accent5, palette.Accent6, palette.Accent1, palette.Accent2 };
        for (int index = 0; index < slots.Length; index++) {
            var slot = colors.Element(a + slots[index])!;
            slot.RemoveNodes();
            slot.Add(new XElement(a + "srgbClr", new XAttribute("val", values[index].Resolve(theme).ToRgbHex())));
        }
        var fonts = elements.Element(a + "fontScheme")!;
        fonts.SetAttributeValue("name", theme.Name);
        fonts.Element(a + "majorFont")!.Element(a + "latin")!.SetAttributeValue("typeface", theme.Fonts.Heading.Family);
        fonts.Element(a + "minorFont")!.Element(a + "latin")!.SetAttributeValue("typeface", theme.Fonts.Body.Family);
        return xml.ToString(SaveOptions.DisableFormatting);
    }
}
