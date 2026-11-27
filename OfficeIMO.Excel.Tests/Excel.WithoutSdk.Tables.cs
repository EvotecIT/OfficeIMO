using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// A loaded workbook lists and reads its tables without opening the Open XML SDK package, and describes them as
    /// the SDK model does.
    /// </summary>
    public class ExcelWithoutSdkTables {
        [Fact]
        public void TableListingsDoNotOpenTheSdkAndMatchTheModel() {
            byte[] source = CreateWorkbook();
            List<string> expected;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                Assert.NotNull(document.OpenXmlDocument);
                expected = Describe(document);
            }
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                Assert.Equal(expected, Describe(document));
                Assert.Empty(document.OpenXmlSdkFallbacks);
            }
        }

        [Fact]
        public void TableRowsAndRangesAreReadWithoutTheSdk() {
            byte[] source = CreateWorkbook();
            using var document = ExcelDocument.Load(new MemoryStream(source));
            ExcelSheet sheet = document.Sheets[1];

            Assert.Equal("B2:D5", sheet.GetTableRange("Orders"));
            Assert.Null(sheet.GetTableRange("Missing"));
            Assert.Equal("B2:D5", sheet.Table("Orders").Range);
            List<Order> rows = sheet.Table("Orders").RowsAs<Order>().ToList();

            Assert.Equal(new[] { "Ann", "Bob" }, rows.Select(row => row.Customer));
            Assert.Equal(new[] { 10d, 20d }, rows.Select(row => row.Amount));
            Assert.Empty(document.OpenXmlSdkFallbacks);
        }

        public sealed class Order {
            public string Customer { get; set; } = string.Empty;
            public double Amount { get; set; }
            public bool Paid { get; set; }
        }

        private static List<string> Describe(ExcelDocument document) {
            var lines = new List<string>();
            foreach (ExcelTableInfo table in document.GetTables()) lines.Add("workbook " + Describe(table));
            foreach (ExcelSheet sheet in document.Sheets) {
                foreach (ExcelTableInfo table in sheet.GetTables()) lines.Add(sheet.Name + " " + Describe(table));
            }
            return lines;
        }

        private static string Describe(ExcelTableInfo table) =>
            string.Join("|", table.Name, table.DisplayName, table.Range, table.SheetName, table.SheetIndex, table.StyleName,
                table.HasHeaderRow, table.TotalsRowShown, table.HasAutoFilter, table.ShowFirstColumn, table.ShowLastColumn,
                table.ShowRowStripes, table.ShowColumnStripes,
                string.Join(",", table.Columns.Select(column => column.Index + ":" + column.Name + ":" + column.TotalsRowFunction)));

        // A sheet without tables, then two tables with different styles, filters and a totals row.
        private static byte[] CreateWorkbook() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.AddSheet("Empty").Cell(1, 1).SetValue("nothing here");
                ExcelSheet data = document.AddSheet("Data");
                data.Cell(2, 2).SetValue("Customer");
                data.Cell(2, 3).SetValue("Amount");
                data.Cell(2, 4).SetValue("Paid");
                data.Cell(3, 2).SetValue("Ann");
                data.Cell(3, 3).SetValue(10d);
                data.Cell(3, 4).SetValue(true);
                data.Cell(4, 2).SetValue("Bob");
                data.Cell(4, 3).SetValue(20d);
                data.Cell(4, 4).SetValue(false);
                data.AddTable("B2:D4", hasHeader: true, name: "Orders", style: ExcelTableStyle.TableStyleLight9)
                    .SetTotals(new Dictionary<string, ExcelTableTotalsFunction> { ["Amount"] = ExcelTableTotalsFunction.Sum });
                data.Cell(8, 1).SetValue("Code");
                data.Cell(8, 2).SetValue("Label");
                data.Cell(9, 1).SetValue(1);
                data.Cell(9, 2).SetValue("one");
                data.AddTable("A8:B9", hasHeader: true, name: "Codes", style: ExcelTableStyle.TableStyleDark2, includeAutoFilter: false)
                    .SetStyle(ExcelTableStyle.TableStyleDark2, showFirstColumn: true, showColumnStripes: true);
                document.Save();
            }
            return stream.ToArray();
        }
    }
}
