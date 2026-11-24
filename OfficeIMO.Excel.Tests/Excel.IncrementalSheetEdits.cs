using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly XNamespace VmlNamespace = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace VmlExcelNamespace = "urn:schemas-microsoft-com:office:excel";

        [Fact]
        public void NotesAddedOneAtATimeKeepOneShapePerCellAcrossSaves() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 3; row++) sheet.SetComment(row, 1, $"First {row}");
                document.Save();

                sheet.SetComment(2, 1, "Replaced");
                sheet.ClearComment(1, 1);
                sheet.SetComment(4, 1, "Fourth");
                sheet.Batch(batch => batch.SetComment(5, 1, "Fifth"));
                sheet.SetComment(3, 1, "Third again");
                document.Save();
            }

            AssertNoteShapes(stream, (2, 1), (3, 1), (4, 1), (5, 1));
            stream.Position = 0;
            using var loaded = ExcelDocument.Load(stream);
            Assert.Equal(
                new[] { "2: Replaced", "3: Third again", "4: Fourth", "5: Fifth" },
                loaded.Sheets[0].GetComments().OrderBy(comment => comment.Row)
                    .Select(comment => $"{comment.Row}: {comment.Text}").ToArray());
        }

        [Fact]
        public void NotesAddedToALoadedWorkbookKeepItsShapesAndTakeNewIds() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.SetComment(1, 1, "One");
                sheet.SetComment(2, 1, "Two");
                document.Save();
            }

            stream.Position = 0;
            int[] existingIds = ReadNoteShapes(stream).Select(shape => shape.Id).ToArray();
            stream.Position = 0;
            using (var document = ExcelDocument.Load(stream)) {
                ExcelSheet sheet = document.Sheets[0];
                sheet.SetComment(3, 1, "Three");
                sheet.SetComment(1, 1, "One again");
                document.Save();
            }

            var shapes = AssertNoteShapes(stream, (1, 1), (2, 1), (3, 1));
            Assert.Contains(shapes, shape => shape.Row == 2 && existingIds.Contains(shape.Id));
            Assert.All(shapes.Where(shape => shape.Row != 2), shape => Assert.True(shape.Id > existingIds.Max()));
        }

        [Fact]
        public void OpenXmlDocumentReadsNotesAddedBeforeAndAfterItWasHandedOut() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.SetComment(1, 1, "Before");

            SpreadsheetDocument package = document.OpenXmlDocument;
            WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal(new[] { (1, 1) }, ReadNoteCells(worksheetPart));

            sheet.SetComment(2, 1, "After");
            Assert.Equal(new[] { (1, 1), (2, 1) }, ReadNoteCells(worksheetPart));
        }

        [Fact]
        public void TablesAddedOneAtATimeRejectOverlapsAndKeepDistinctParts() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                AddKeyValueTable(sheet, 1, "First");
                AddKeyValueTable(sheet, 4, "Second");
                document.Save();

                Assert.Throws<InvalidOperationException>(() => sheet.AddTable("B2:C4", true, "Overlap"));
                AddKeyValueTable(sheet, 7, "Third");
                // The header of the table just added is owned by its schema.
                sheet.Cell(7, 2).SetValue("Amount");
                Assert.Throws<InvalidOperationException>(() => sheet.AddTable("A8:B9", true, "Overlap"));
                document.Save();
            }

            stream.Position = 0;
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                TablePart[] tableParts = worksheetPart.Worksheet.GetFirstChild<TableParts>()!.Elements<TablePart>().ToArray();
                Assert.Equal(3, tableParts.Length);
                Assert.Equal(3U, worksheetPart.Worksheet.GetFirstChild<TableParts>()!.Count!.Value);
                Table[] tables = tableParts.Select(part => ((TableDefinitionPart)worksheetPart.GetPartById(part.Id!.Value!)).Table).ToArray();
                Assert.Equal(new[] { "First", "Second", "Third" }, tables.Select(table => table.Name!.Value).ToArray());
                Assert.Equal(3, tables.Select(table => table.Id!.Value).Distinct().Count());
                Assert.Equal("Amount", tables[2].TableColumns!.Elements<TableColumn>().ElementAt(1).Name!.Value);
            }
        }

        [Fact]
        public void TablesAddedToALoadedWorkbookSeeItsTablesAndResizes() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                AddKeyValueTable(document.AddSheet("Data"), 1, "Loaded");
                document.Save();
            }

            stream.Position = 0;
            using (var document = ExcelDocument.Load(stream)) {
                ExcelSheet sheet = document.Sheets[0];
                Assert.Throws<InvalidOperationException>(() => sheet.AddTable("B2:C3", true, "Overlap"));
                sheet.ResizeTable("Loaded", "A1:B6");
                Assert.Throws<InvalidOperationException>(() => sheet.AddTable("A6:B7", true, "Overlap"));
                AddKeyValueTable(sheet, 8, "Added");
                document.Save();
            }

            stream.Position = 0;
            using var reloaded = ExcelDocument.Load(stream);
            Assert.Equal(
                new[] { "Loaded A1:B6", "Added A8:B9" },
                reloaded.Sheets[0].GetTables().Select(table => $"{table.Name} {table.Range}").ToArray());
        }

        [Fact]
        public void NamesSetBetweenCellWritesMarkExistingFormulasForRecalculation() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 2).SetFormula("A1*2");
                for (int i = 1; i <= 4; i++) {
                    sheet.Cell(i, 1).SetValue(i);
                    document.SetNamedRange($"Name{i}", $"'Data'!$A${i}", scope: null, save: false);
                    if (i == 2) document.Save();
                }
                document.SetNamedRange("Name1", "'Data'!$A$4", scope: null, save: false);
                document.Save();
            }

            stream.Position = 0;
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false)) {
                CellFormula formula = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<CellFormula>().Single();
                Assert.True(formula.CalculateCell?.Value);
            }

            stream.Position = 0;
            using var loaded = ExcelDocument.Load(stream);
            IReadOnlyDictionary<string, string> names = loaded.GetAllNamedRanges();
            Assert.Equal(4, names.Count);
            Assert.Equal("'Data'!$A$4", names["Name1"]);
        }

        private static void AddKeyValueTable(ExcelSheet sheet, int top, string name) {
            sheet.Cell(top, 1).SetValue("Key");
            sheet.Cell(top, 2).SetValue("Value");
            sheet.Cell(top + 1, 1).SetValue(top);
            sheet.Cell(top + 1, 2).SetValue("v");
            sheet.AddTable($"A{top}:B{top + 1}", true, name);
        }

        private static List<(int Row, int Column, int Id)> AssertNoteShapes(Stream stream, params (int Row, int Column)[] cells) {
            stream.Position = 0;
            List<(int Row, int Column, int Id)> shapes = ReadNoteShapes(stream);
            Assert.Equal(cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray(),
                shapes.Select(shape => (shape.Row, shape.Column)).OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray());
            Assert.Equal(shapes.Count, shapes.Select(shape => shape.Id).Distinct().Count());
            return shapes;
        }

        private static List<(int Row, int Column, int Id)> ReadNoteShapes(Stream stream) {
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            return ReadNoteShapes(package.WorkbookPart!.WorksheetParts.Single());
        }

        private static (int Row, int Column)[] ReadNoteCells(WorksheetPart worksheetPart)
            => ReadNoteShapes(worksheetPart).Select(shape => (shape.Row, shape.Column)).ToArray();

        private static List<(int Row, int Column, int Id)> ReadNoteShapes(WorksheetPart worksheetPart) {
            string id = worksheetPart.Worksheet.GetFirstChild<LegacyDrawing>()!.Id!.Value!;
            using Stream vml = worksheetPart.GetPartById(id).GetStream(FileMode.Open, FileAccess.Read);
            XDocument document = XDocument.Load(vml);
            return document.Root!.Elements(VmlNamespace + "shape").Select(shape => {
                XElement data = shape.Element(VmlExcelNamespace + "ClientData")!;
                return (
                    int.Parse(data.Element(VmlExcelNamespace + "Row")!.Value, CultureInfo.InvariantCulture) + 1,
                    int.Parse(data.Element(VmlExcelNamespace + "Column")!.Value, CultureInfo.InvariantCulture) + 1,
                    int.Parse(shape.Attribute("id")!.Value.Substring("_x0000_s".Length), CultureInfo.InvariantCulture));
            }).ToList();
        }
    }
}
