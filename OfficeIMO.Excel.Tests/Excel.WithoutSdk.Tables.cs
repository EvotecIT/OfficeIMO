using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
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

        [Fact]
        public void AWorksheetWithoutASheetEntryListsItsTablesAsTheModelDoes() {
            // The workbook still relates the Data worksheet, but no sheet entry names it.
            byte[] source = Rewrite(CreateWorkbook(), "xl/workbook.xml", xml => Regex.Replace(xml, "<(\\w+:)?sheet [^>]*name=\"Data\"[^>]*/>", string.Empty));
            List<string> listed = AssertListingsMatchTheModel(source);
            Assert.Contains(listed, line => line.StartsWith("workbook Orders|Orders|B2:D5||-1|", StringComparison.Ordinal));
        }

        [Fact]
        public void RowCountsWithSurroundingWhitespaceReadAsTheModelReadsThem() {
            byte[] source = RewriteTables(CreateWorkbook(), xml => xml.Replace("totalsRowCount=\"1\"", "totalsRowCount=\" 1 \""));
            AssertListingsMatchTheModel(source);
            using var document = ExcelDocument.Load(new MemoryStream(source));
            Assert.Equal(new[] { "Ann", "Bob" }, document.Sheets[1].Table("Orders").RowsAs<Order>().Select(row => row.Customer));
        }

        [Fact]
        public void ADamagedTableOnOneWorksheetDoesNotHideTablesOnAnother() {
            byte[] source = CreateWorkbook(sumSheet: true);
            // Damage the table part of the Other sheet; the formula on the Data sheet reads only its own table.
            source = RewriteTables(source, xml => xml.Contains("name=\"Others\"") ? xml.Substring(0, xml.Length / 2) : xml);
            using var document = ExcelDocument.Load(new MemoryStream(source));
            document.Calculate();
            Assert.True(document.Sheets[1].TryGetCachedFormulaValue(12, 1, out string? total));
            Assert.Equal("30", total);
        }

        private static List<string> AssertListingsMatchTheModel(byte[] source) {
            List<string> expected;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                Assert.NotNull(document.OpenXmlDocument);
                expected = Describe(document);
            }
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                Assert.Equal(expected, Describe(document));
            }
            return expected;
        }

        private static byte[] RewriteTables(byte[] package, Func<string, string> rewrite) {
            foreach (string name in EntryNames(package).Where(name => name.StartsWith("xl/tables/", StringComparison.Ordinal))) {
                package = Rewrite(package, name, rewrite);
            }
            return package;
        }

        private static List<string> EntryNames(byte[] package) {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            return archive.Entries.Select(entry => entry.FullName).ToList();
        }

        private static byte[] Rewrite(byte[] package, string entryName, Func<string, string> rewrite) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = archive.GetEntry(entryName)!;
                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
                entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(entryName).Open(), new UTF8Encoding(false));
                writer.Write(rewrite(text));
            }
            return stream.ToArray();
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
        private static byte[] CreateWorkbook(bool sumSheet = false) {
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
                if (sumSheet) {
                    data.Cell(12, 1).SetFormula("SUM(Orders[Amount])");
                    ExcelSheet other = document.AddSheet("Other");
                    other.Cell(1, 1).SetValue("Key");
                    other.Cell(2, 1).SetValue("k");
                    other.AddTable("A1:A2", hasHeader: true, name: "Others", style: ExcelTableStyle.TableStyleLight1);
                }
                document.Save();
            }
            return stream.ToArray();
        }
    }
}
