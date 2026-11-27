using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Text written to a table header of a loaded workbook renames the column without opening the Open XML SDK
    /// package, and saves the workbook the model saves.
    /// </summary>
    public class ExcelWithoutSdkTableHeaders {
        [Fact]
        public void RenamingAColumnSavesWhatTheModelSaves() {
            byte[] source = CreateWorkbook(addCalculationChain: true);
            AssertSavesAsTheModel(source, sheet => sheet.Cell(2, 3).SetValue("  Total "), expectOwned: true, out string saved);
            Assert.Contains("Orders|B2:D4|1:Customer,2:Total,3:Paid", saved);
            Assert.Contains("calcPr calcId=191029 fullCalcOnLoad=1 forceFullCalc=1", saved);
            Assert.Contains("calcChain False", saved);
        }

        [Fact]
        public void ARenamedColumnIsListedWithoutTheSdkAndReadByItsNewName() {
            using var document = ExcelDocument.Load(new MemoryStream(CreateWorkbook()));
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(2, 3).SetValue("Total");

            Assert.Equal(new[] { "Customer", "Total", "Paid" }, sheet.GetTables()[0].Columns.Select(column => column.Name));
            Assert.Empty(document.OpenXmlSdkFallbacks);
            // Cell reads after recorded changes go through the model, which receives the renamed table.
            Assert.Equal(new[] { 10d, 20d }, sheet.Table("Orders").RowsAs<Renamed>().Select(row => row.Total));
        }

        public sealed class Renamed {
            public string Customer { get; set; } = string.Empty;
            public double Total { get; set; }
        }

        [Fact]
        public void WritingAColumnsOwnNameChangesOnlyTheCell() =>
            AssertSavesAsTheModel(CreateWorkbook(), sheet => sheet.Cell(2, 2).SetValue("Customer"), expectOwned: true, out _);

        [Fact]
        public void AHeaderStoredAsANumberIsRewrittenAsItsColumnName() {
            byte[] source = CreateWorkbook(numericHeader: true);
            AssertSavesAsTheModel(source, sheet => sheet.Cell(2, 2).SetValue("Client"), expectOwned: true, out string saved);
            Assert.Contains("D2 s 2024", saved);
        }

        [Fact]
        public void AHeaderHeldAsInlineTextKeepsItWhenItAlreadyMatches() {
            byte[] source = CreateWorkbook(inlineHeader: true);
            AssertSavesAsTheModel(source, sheet => sheet.Cell(2, 2).SetValue("Client"), expectOwned: true, out string saved);
            Assert.Contains("D2 inline Paid", saved);
        }

        [Fact]
        public void ADuplicateColumnNameFailsAsItDoesThroughTheModel() {
            byte[] source = CreateWorkbook();
            foreach (bool openSdkFirst in new[] { true, false }) {
                using var document = ExcelDocument.Load(new MemoryStream(source));
                if (openSdkFirst) document.OpenDeferredSdk("comparison");
                ArgumentException error = Assert.Throws<ArgumentException>(() => document.Sheets[0].Cell(2, 3).SetValue("customer"));
                Assert.Contains("duplicated", error.Message);
                Assert.Equal(new[] { "Customer", "Amount", "Paid" }, document.Sheets[0].GetTables()[0].Columns.Select(column => column.Name));
            }
        }

        [Fact]
        public void ARenameFormulasReferToGoesThroughTheModel() {
            byte[] source = CreateWorkbook(structuredReference: true);
            AssertSavesAsTheModel(source, sheet => sheet.Cell(2, 3).SetValue("Total"), expectOwned: false, out string saved);
            Assert.Contains("SUM(Orders[Total])", saved);
        }

        [Fact]
        public void AReferenceSpelledWithCharacterReferencesGoesThroughTheModel() {
            // The worksheet spells the column in the formula with a character reference, so its bytes never hold the name.
            byte[] source = RewriteWorksheet(CreateWorkbook(structuredReference: true), xml => xml.Replace("Orders[Amount]", "Orders[&#65;mount]"));
            AssertSavesAsTheModel(source, sheet => sheet.Cell(2, 3).SetValue("Total"), expectOwned: false, out string saved);
            Assert.Contains("SUM(Orders[Total])", saved);
        }

        private static byte[] RewriteWorksheet(byte[] package, Func<string, string> rewrite) {
            using var stream = new MemoryStream();
            stream.Write(package, 0, package.Length);
            using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true)) {
                System.IO.Compression.ZipArchiveEntry entry = archive.Entries.Single(item => item.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal));
                string name = entry.FullName, text;
                using (var reader = new StreamReader(entry.Open())) text = reader.ReadToEnd();
                string rewritten = rewrite(text);
                Assert.NotEqual(text, rewritten);
                entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
                writer.Write(rewritten);
            }
            return stream.ToArray();
        }

        private static void AssertSavesAsTheModel(byte[] source, Action<ExcelSheet> edit, bool expectOwned, out string saved) {
            string expected = Save(source, edit, openSdkFirst: true, out _);
            saved = Save(source, edit, openSdkFirst: false, out bool owned);
            Assert.Equal(expected, saved);
            Assert.Equal(expectOwned, owned);
        }

        private static string Save(byte[] source, Action<ExcelSheet> edit, bool openSdkFirst, out bool owned) {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                // The model, opened without handing the package out, which would make the save repair it.
                if (openSdkFirst) document.OpenDeferredSdk("comparison");
                edit(document.Sheets[0]);
                document.Save(output);
                // The save may still hand the workbook to the model, so it counts too.
                owned = !openSdkFirst && document.OpenXmlSdkFallbacks.Count == 0;
            }
            return Describe(output.ToArray());
        }

        // What the edit can change: table columns, calculation properties and chain, and every cell's type, text and formula.
        private static string Describe(byte[] package) {
            using var stream = new MemoryStream(package);
            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbook = document.WorkbookPart!;
            var lines = new List<string>();
            SharedStringItem[] shared = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
            foreach (WorksheetPart worksheet in workbook.WorksheetParts) {
                foreach (TableDefinitionPart part in worksheet.TableDefinitionParts) {
                    Table table = part.Table!;
                    lines.Add(table.Name + "|" + table.Reference + "|" + string.Join(",", table.TableColumns!.Elements<TableColumn>()
                        .Select(column => column.Id + ":" + column.Name)) + " count=" + table.TableColumns!.Count);
                }
                foreach (Cell cell in worksheet.Worksheet.Descendants<Cell>()) {
                    string type = cell.DataType?.Value is { } kind ? (kind == CellValues.SharedString ? "s" : kind == CellValues.InlineString ? "inline" : kind.ToString()) : "n";
                    string text = cell.DataType?.Value == CellValues.SharedString ? shared[int.Parse(cell.CellValue!.Text)].InnerText
                        : cell.DataType?.Value == CellValues.InlineString ? cell.InlineString!.InnerText : cell.CellValue?.Text ?? string.Empty;
                    lines.Add(cell.CellReference + " " + type + " " + text + (cell.CellFormula != null ? " =" + cell.CellFormula.Text : string.Empty));
                }
            }
            CalculationProperties? calc = workbook.Workbook.CalculationProperties;
            lines.Add(calc == null ? "no calcPr" : "calcPr " + string.Join(" ", calc.GetAttributes().Select(attribute => attribute.LocalName + "=" + attribute.Value)));
            lines.Add("calcChain " + (workbook.CalculationChainPart != null));
            return string.Join("\n", lines);
        }

        private static byte[] CreateWorkbook(bool addCalculationChain = false, bool numericHeader = false, bool structuredReference = false, bool inlineHeader = false) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
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
                data.AddTable("B2:D4", hasHeader: true, name: "Orders", style: ExcelTableStyle.TableStyleLight9);
                data.Cell(7, 2).SetFormula(structuredReference ? "SUM(Orders[Amount])" : "SUM(C3:C4)");
                document.Save();
            }
            if (!addCalculationChain && !numericHeader && !inlineHeader) return stream.ToArray();
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, true)) {
                WorkbookPart workbook = package.WorkbookPart!;
                if (addCalculationChain) {
                    CalculationChainPart chain = workbook.AddNewPart<CalculationChainPart>();
                    chain.CalculationChain = new CalculationChain(new CalculationCell { CellReference = "B7", SheetId = 1 });
                }
                if (inlineHeader) {
                    Cell header = workbook.WorksheetParts.First().Worksheet.Descendants<Cell>().First(cell => cell.CellReference == "D2");
                    header.DataType = CellValues.InlineString;
                    header.CellValue = null;
                    header.InlineString = new InlineString(new Text("Paid"));
                }
                if (numericHeader) {
                    // The column is named 2024, and its header cell holds that number rather than text.
                    WorksheetPart worksheet = workbook.WorksheetParts.First();
                    TableDefinitionPart table = worksheet.TableDefinitionParts.First();
                    table.Table!.TableColumns!.Elements<TableColumn>().ElementAt(2).Name = "2024";
                    Cell header = worksheet.Worksheet.Descendants<Cell>().First(cell => cell.CellReference == "D2");
                    header.DataType = null;
                    header.CellValue = new CellValue("2024");
                }
            }
            return stream.ToArray();
        }
    }
}
