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
        internal static async Task<int> WriteStreamingWorksheet<T>(ZipArchive archive, DirectDataSetWorkbookModel model,
            IAsyncEnumerable<T> rows, Action<ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken) {
            var stylePlan = DirectStylePlan.Create(model);
            var columns = CreateColumnWritePlans(model, stylePlan, cancellationToken);
            var sheet = model.Sheets[0];
            int count = await WriteAsyncWorksheetRows(archive, sheet, model.DateTimeOffsetWriteStrategy,
                model.DateSystem, columns[0], rows, writeRow, ExcelA1.MaxRows - (sheet.IncludeHeaders ? 1 : 0),
                null, stylePlan.TabularStyles, cancellationToken).ConfigureAwait(false);
            if (!sheet.HasTable) return count;
            string index = InvariantNumberText.Get(sheet.Index);
            WriteTextEntry(archive, "xl/worksheets/_rels/sheet" + index + ".xml.rels",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table" + index + ".xml\"/></Relationships>");
            WriteTable(archive, sheet, ExcelSheet.BuildObjectExportRange(1, sheet.Table.ColumnCount, count, sheet.IncludeHeaders));
            return count;
        }

        /// <summary>Writes bounded workbook metadata only after all worksheet streams have completed.</summary>
        internal static void CompleteStreamingWorkbook(ZipArchive archive, DirectDataSetWorkbookModel model, OfficeTheme theme) {
            WritePackagePreamble(archive, model, DirectStylePlan.Create(model), false, theme);
            WriteTextEntry(archive, "xl/theme/theme1.xml", CreateWorkbookThemeXml(theme));
        }


    }
}
