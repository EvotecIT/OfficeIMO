using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreservesExternalPivotWorksheetSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            PivotTableCacheDefinitionPart cachePart = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!;
            WorksheetSource source = cachePart.PivotCacheDefinition!
                .CacheSource!.WorksheetSource!;
            source.Id = "rIdExternal";
            source.Sheet = "Data";
            source.Reference = $"A{A1.MaxRows}";
            cachePart.PivotCacheDefinition.RefreshOnLoad = false;
            cachePart.PivotCacheDefinition.SaveData = true;

            sheet.InsertRows(1);

            Assert.Equal($"A{A1.MaxRows}", source.Reference!.Value);
            Assert.False(cachePart.PivotCacheDefinition.RefreshOnLoad!.Value);
            Assert.True(cachePart.PivotCacheDefinition.SaveData!.Value);
        }

        [Fact]
        public void Test_StructuralRows_PreservesSameRowTableFormulaAtFirstDataRow() {
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
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
            TableColumn resultColumn = table.Descendants<TableColumn>()
                .Single(column => column.Name?.Value == "Result");
            var formula = new CalculatedColumnFormula("A2*2");
            resultColumn.Append(formula);

            sheet.InsertRows(2);

            Assert.Equal("A1:B4", table.Reference!.Value);
            Assert.Equal("A2*2", formula.Text);
        }

        [Fact]
        public void Test_StructuralRows_RequestsRecalculationForDataTableOnlyWorkbook() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue(1);
            sheet.CellAt(2, 2).SetValue(2);
            Cell owner = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B2");
            owner.CellFormula = new CellFormula {
                FormulaType = CellFormulaValues.DataTable,
                Reference = "B2:B3",
                R1 = "A1"
            };
            document.WorkbookRoot.Append(new CalculationProperties {
                CalculationMode = CalculateModeValues.Manual,
                FullCalculationOnLoad = false,
                ForceFullCalculation = false
            });

            sheet.InsertRows(1);

            CalculationProperties calculation =
                document.WorkbookRoot.GetFirstChild<CalculationProperties>()!;
            Assert.True(calculation.FullCalculationOnLoad!.Value);
            Assert.True(calculation.ForceFullCalculation!.Value);
        }
    }
}
