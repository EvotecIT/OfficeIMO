using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xnsv = DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreservesRelativeTableFormulaOffsetsAtFirstDataRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Value");
            sheet.CellAt(1, 2).SetValue("Result");
            sheet.CellAt(2, 1).SetValue(10);
            sheet.CellAt(3, 1).SetValue(20);
            sheet.AddTable(
                "A1:B3",
                hasHeader: true,
                name: "CalculatedData",
                OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
            TableColumn resultColumn = table.Descendants<TableColumn>()
                .Single(column => column.Name?.Value == "Result");
            var formula = new CalculatedColumnFormula("A3+SUM(A3:A4)+$A$3");
            resultColumn.Append(formula);

            sheet.InsertRows(2);

            Assert.Equal("A1:B4", table.Reference!.Value);
            Assert.Equal("A3+SUM(A3:A4)+$A$4", formula.Text);

            sheet.DeleteRows(2);

            Assert.Equal("A1:B3", table.Reference!.Value);
            Assert.Equal("A3+SUM(A3:A4)+$A$3", formula.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsNamedSheetViewFilters() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Xnsv.NsvFilter filter = AddNamedSheetViewFilter(sheet, "A5:C10");

            sheet.InsertRows(5, 2);

            Assert.Equal("A7:C12", filter.Ref!.Value);

            sheet.DeleteRows(7, 2);

            Assert.Equal("A7:C10", filter.Ref!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsNamedSheetViewFilters() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Keep");
            Xnsv.NsvFilter filter = AddNamedSheetViewFilter(
                sheet,
                $"A{A1.MaxRows}:C{A1.MaxRows}");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}:C{A1.MaxRows}", filter.Ref!.Value);
            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RejectsDeletionOfNamedPivotSourceHeader() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            WorksheetSource source = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            source.Reference = null;
            source.Sheet = null;
            source.Name = "PivotSource";
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                new DefinedName("'Data'!$A$1:$B$3") { Name = "PivotSource" });

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(1));

            Assert.Contains("header row", exception.Message);
            Assert.Equal("PivotSource", source.Name!.Value);
            Assert.Equal("Region", sheet.CellAt(1, 1).GetValue<string>());
        }

        private static Xnsv.NsvFilter AddNamedSheetViewFilter(
            ExcelSheet sheet,
            string reference) {
            var filter = new Xnsv.NsvFilter {
                FilterId = "{956D1A31-7677-46F9-BA3B-C56A2CA5EB48}",
                Ref = reference
            };
            var view = new Xnsv.NamedSheetView(filter) {
                Name = "Saved filter",
                Id = "{82030999-1976-4F4D-A705-9218D481B69D}"
            };
            NamedSheetViewsPart part = sheet.WorksheetPart.AddNewPart<NamedSheetViewsPart>();
            part.NamedSheetViews = new Xnsv.NamedSheetViews(view);
            part.NamedSheetViews.Save();
            return filter;
        }
    }
}
