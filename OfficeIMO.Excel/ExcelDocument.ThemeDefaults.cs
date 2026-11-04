using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Creates the default font used by both editable and streaming workbook writers.</summary>
        internal static Font CreateDefaultFont() => new Font(
            new FontSize { Val = OfficeTheme.Default.TypeScale.Body.ToPoints() },
            new FontName { Val = OfficeTheme.Default.Fonts.Body.Family });

        /// <summary>Retains the Office effect and script-font schemes while applying the shared visual identity.</summary>
        private static byte[] CreateDefaultThemeBytes() {
            OfficeTheme theme = OfficeTheme.Default;
            XDocument document = XDocument.Parse(Encoding.UTF8.GetString(
                LoadEmbeddedResource("OfficeIMO.Excel.Resources.theme1.xml")));
            XNamespace ns = DrawingNamespace;
            XElement root = document.Root!;
            XElement elements = root.Element(ns + "themeElements")!;
            XElement colors = elements.Element(ns + "clrScheme")!;
            root.SetAttributeValue("name", "OfficeIMO " + theme.Name);
            colors.SetAttributeValue("name", theme.Name);
            void SetColor(string slot, OfficeColor value) => colors.Element(ns + slot)!.ReplaceNodes(
                new XElement(ns + "srgbClr", new XAttribute("val", value.Resolve(theme).ToRgbHex())));
            SetColor("dk1", theme.Palette.Text);
            SetColor("lt1", OfficeColor.White);
            SetColor("dk2", theme.Palette.TextMuted);
            SetColor("lt2", theme.Surfaces.Surface);
            SetColor("accent1", theme.Palette.Accent1);
            SetColor("accent2", theme.Palette.Accent2);
            SetColor("accent3", theme.Palette.Accent3);
            SetColor("accent4", theme.Palette.Accent4);
            SetColor("accent5", theme.Palette.Accent5);
            SetColor("accent6", theme.Palette.Accent6);
            SetColor("hlink", theme.TextStyle.LinkColor);
            SetColor("folHlink", theme.Palette.Accent4);
            XElement fonts = elements.Element(ns + "fontScheme")!;
            fonts.SetAttributeValue("name", theme.Name);
            fonts.Element(ns + "majorFont")!.Element(ns + "latin")!.SetAttributeValue("typeface", theme.Fonts.Heading.Family);
            fonts.Element(ns + "minorFont")!.Element(ns + "latin")!.SetAttributeValue("typeface", theme.Fonts.Body.Family);
            return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        }
    }
}
