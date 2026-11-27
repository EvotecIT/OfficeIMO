using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// A loaded workbook calculates over its own worksheet XML and saves the results without opening the Open XML SDK
    /// package. The saved cells must match a calculation through the SDK model.
    /// </summary>
    public class ExcelWithoutSdkCalculation {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        [Fact]
        public void CalculatingAndSavingALoadedWorkbookDoesNotOpenTheSdk() {
            byte[] source = CreateWorkbook();
            using var document = ExcelDocument.Load(new MemoryStream(source));
            document.Sheets[0].Cell(2, 1).SetValue(5d);

            Assert.Equal(8, document.Calculate());
            Assert.True(document.Sheets[0].TryGetCachedFormulaValue(2, 2, out string? doubled));
            Assert.Equal("10", doubled);
            using var output = new MemoryStream();
            document.Save(output);

            Assert.Empty(document.OpenXmlSdkFallbacks);
        }

        [Fact]
        public void CalculatedCellsMatchACalculationThroughTheSdk() {
            byte[] source = CreateWorkbook();
            Assert.Equal(CellsOf(Calculate(source, openSdkFirst: true)), CellsOf(Calculate(source, openSdkFirst: false)));
        }

        [Fact]
        public void ResultsCalculatedBeforeTheSdkOpensAreKept() {
            byte[] source = CreateWorkbook();
            byte[] expected = Calculate(source, openSdkFirst: true);
            byte[] actual;
            using (var document = ExcelDocument.Load(new MemoryStream(source))) {
                document.Sheets[0].Cell(2, 1).SetValue(5d);
                document.Calculate();
                Assert.NotNull(document.OpenXmlDocument);
                using var output = new MemoryStream();
                document.Save(output);
                actual = output.ToArray();
            }
            Assert.Equal(CellsOf(expected), CellsOf(actual));
        }

        [Fact]
        public void ASecondCalculationAfterASaveStartsFromTheSavedResults() {
            byte[] source = CreateWorkbook();
            using var document = ExcelDocument.Load(new MemoryStream(source));
            document.Calculate();
            document.Save(new MemoryStream());
            document.Sheets[0].Cell(3, 1).SetValue(10d);
            document.Calculate();

            Assert.True(document.Sheets[0].TryGetCachedFormulaValue(1, 6, out string? sum));
            Assert.Equal("13", sum);
            Assert.Empty(document.OpenXmlSdkFallbacks);
        }

        [Fact]
        public void ASheetWithAnUnnumberedRowCalculatesAsTheModelDoes() {
            // The master names B5 in a row that omits its number; the dependent after it has no reference of its own.
            // The model places these cells by their row and writes results into them, so the sheet calculates through it.
            byte[] source = Package(new Row(
                FormulaCell("B5", new CellFormula("A6*2") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U, Reference = "B5:C5" }, "0"),
                new Cell { CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U }, CellValue = new CellValue("0") }),
                new Row(Number("A6", "4"), Number("B6", "5")) { RowIndex = 6U });
            Assert.Equal(CellsOf(Calculate(source, openSdkFirst: true, edit: false)), CellsOf(Calculate(source, openSdkFirst: false, edit: false)));
        }

        [Fact]
        public void ARichErrorFromTheModelIsNotOverwrittenByTheCalculateFlag() {
            // SEQUENCE(0) is #CALC!, a rich error, which opens the SDK model; the flags written after it must keep it.
            byte[] source = Package(new Row(
                FormulaCell("B1", new CellFormula("SEQUENCE(0)") { FormulaType = CellFormulaValues.Array, Reference = "B1", CalculateCell = true }, "7")) { RowIndex = 1U });
            byte[] expected = Calculate(source, openSdkFirst: true, edit: false);
            byte[] actual = Calculate(source, openSdkFirst: false, edit: false);
            Assert.Equal(CellsOf(expected), CellsOf(actual));
            Assert.DoesNotContain(CellsOf(actual), cell => cell.StartsWith("B1|", StringComparison.Ordinal) && cell.Contains("|7|"));
        }

        [Fact]
        public void AnArrayMemberWrittenAfterARichErrorGoesToTheModel() {
            // B2 takes #CALC! from A2, which opens the SDK model; B3 changes afterwards and has no formula of its own.
            byte[] source = Package(
                new Row(Number("A1", "1"), FormulaCell("B1", new CellFormula("A1:A3") { FormulaType = CellFormulaValues.Array, Reference = "B1:B3" }, "1")) { RowIndex = 1U },
                new Row(FormulaCell("A2", new CellFormula("SEQUENCE(0)") { FormulaType = CellFormulaValues.Array, Reference = "A2" }, "2"), Number("B2", "2")) { RowIndex = 2U },
                new Row(Number("A3", "30"), Number("B3", "3")) { RowIndex = 3U });
            Assert.Equal(CellsOf(Calculate(source, openSdkFirst: true, edit: false)), CellsOf(Calculate(source, openSdkFirst: false, edit: false)));
        }

        [Fact]
        public void InspectingFormulasReadsTheSheetsAfterOneThatOpensTheSdk() {
            // The Lookups sheet reads saved pivot tables, which opens the SDK model part way through the workbook.
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", "aggregation-conformance.xlsx");
            List<string> expected;
            using (var document = ExcelDocument.Load(path)) {
                Assert.NotNull(document.OpenXmlDocument);
                expected = document.InspectFormulas().Formulas.Select(cell => cell.SheetName + "!" + cell.CellReference + "=" + cell.CachedValue).ToList();
            }
            using (var document = ExcelDocument.Load(path)) {
                Assert.Equal(expected, document.InspectFormulas().Formulas.Select(cell => cell.SheetName + "!" + cell.CellReference + "=" + cell.CachedValue));
            }
        }

        private static byte[] Package(params Row[] rows) {
            using var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbook = package.AddWorkbookPart();
                workbook.Workbook = new Workbook(new Sheets(new Sheet { Name = "Data", SheetId = 1U, Id = "rId1" }));
                workbook.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet(new SheetData(rows));
            }
            return stream.ToArray();
        }

        private static byte[] Calculate(byte[] source, bool openSdkFirst, bool edit) {
            using var document = ExcelDocument.Load(new MemoryStream(source));
            if (openSdkFirst) Assert.NotNull(document.OpenXmlDocument);
            if (edit) document.Sheets[0].Cell(2, 1).SetValue(5d);
            document.Calculate();
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }

        private static byte[] Calculate(byte[] source, bool openSdkFirst) {
            using var document = ExcelDocument.Load(new MemoryStream(source));
            if (openSdkFirst) Assert.NotNull(document.OpenXmlDocument);
            document.Sheets[0].Cell(2, 1).SetValue(5d);
            document.Calculate();
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }

        // Numbers, a shared formula block, text, a boolean, an error, a flagged sum and a fixed array over a cell
        // that does not exist yet.
        private static byte[] CreateWorkbook() {
            using var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbook = package.AddWorkbookPart();
                workbook.Workbook = new Workbook(new Sheets(new Sheet { Name = "Data", SheetId = 1U, Id = "rId1" }));
                WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>("rId1");
                SharedStringTablePart strings = workbook.AddNewPart<SharedStringTablePart>();
                strings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("label")));
                worksheet.Worksheet = new Worksheet(new SheetData(
                    new Row(
                        Number("A1", "1"),
                        FormulaCell("B1", new CellFormula("A1*2") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U, Reference = "B1:B3" }, "2"),
                        FormulaCell("C1", new CellFormula("\"x\"&A1"), "x1", CellValues.String),
                        FormulaCell("D1", new CellFormula("A1>1"), "0", CellValues.Boolean),
                        FormulaCell("E1", new CellFormula("1/0"), null),
                        FormulaCell("F1", new CellFormula("SUM(A1:A3)") { CalculateCell = true }, "6"),
                        new Cell { CellReference = "G1", DataType = CellValues.SharedString, CellValue = new CellValue("0") }) { RowIndex = 1U },
                    new Row(
                        Number("A2", "2"),
                        FormulaCell("B2", new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U }, "4")) { RowIndex = 2U },
                    new Row(
                        Number("A3", "3"),
                        FormulaCell("B3", new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U }, "6")) { RowIndex = 3U },
                    new Row(
                        FormulaCell("B5", new CellFormula("A1:A2") { FormulaType = CellFormulaValues.Array, Reference = "B5:B6" }, "1")) { RowIndex = 5U }));
            }
            return stream.ToArray();
        }

        private static Cell Number(string reference, string value) => new Cell { CellReference = reference, CellValue = new CellValue(value) };

        private static Cell FormulaCell(string reference, CellFormula formula, string? value, CellValues? type = null) {
            var cell = new Cell { CellReference = reference, CellFormula = formula };
            if (value != null) cell.CellValue = new CellValue(value);
            if (type != null) cell.DataType = type.Value;
            return cell;
        }

        /// <summary>Each cell of the first worksheet: its reference, type, value, formula text and formula attributes.</summary>
        private static List<string> CellsOf(byte[] package) {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            ZipArchiveEntry entry = archive.Entries.First(item => item.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal));
            using Stream stream = entry.Open();
            XDocument xml = XDocument.Load(stream);
            var cells = new List<string>();
            foreach (XElement cell in xml.Descendants(Main + "c")) {
                XElement? formula = cell.Element(Main + "f");
                string formulaText = formula == null ? string.Empty
                    : formula.Value + "|" + string.Join(",", formula.Attributes().OrderBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
                        .Select(attribute => attribute.Name.LocalName + "=" + NormalizeBoolean(attribute)));
                cells.Add((string?)cell.Attribute("r") + "|" + (string?)cell.Attribute("t") + "|" + (string?)cell.Element(Main + "v") + "|" + formulaText);
            }
            return cells;
        }

        private static string NormalizeBoolean(XAttribute attribute) =>
            attribute.Name.LocalName == "ca" ? (attribute.Value is "1" or "true" ? "1" : "0") : attribute.Value;
    }
}
