using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private static partial class DirectDataSetWorkbookWriter {
        /// <summary>Streams one worksheet, then emits its table with the final observed row count.</summary>
        internal static async Task WriteStreamingWorksheet<T>(ZipArchive archive, DirectDataSetWorkbookModel model,
            IAsyncEnumerable<T> rows, Action<ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken) {
            var stylePlan = DirectStylePlan.Create(model);
            var columns = CreateColumnWritePlans(model, stylePlan, cancellationToken);
            var sheet = model.Sheets[0];
            int count = await WriteAsyncWorksheetRows(archive, sheet, model.DateTimeOffsetWriteStrategy,
                model.DateSystem, columns[0], rows, writeRow, ExcelA1.MaxRows - (sheet.IncludeHeaders ? 1 : 0),
                null, stylePlan.TabularStyles, cancellationToken).ConfigureAwait(false);
            if (!sheet.HasTable) return;
            string index = InvariantNumberText.Get(sheet.Index);
            WriteTextEntry(archive, "xl/worksheets/_rels/sheet" + index + ".xml.rels",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table" + index + ".xml\"/></Relationships>");
            WriteTable(archive, sheet, ExcelSheet.BuildObjectExportRange(1, sheet.Table.ColumnCount, count, sheet.IncludeHeaders));
        }

        /// <summary>Writes bounded workbook metadata only after all worksheet streams have completed.</summary>
        internal static void CompleteStreamingWorkbook(ZipArchive archive, DirectDataSetWorkbookModel model, OfficeTheme theme) {
            WritePackagePreamble(archive, model, DirectStylePlan.Create(model), false, theme);
            WriteTextEntry(archive, "xl/theme/theme1.xml", CreateStreamingThemeXml(theme));
        }

        private static string CreateStreamingThemeXml(OfficeTheme theme) {
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
}
