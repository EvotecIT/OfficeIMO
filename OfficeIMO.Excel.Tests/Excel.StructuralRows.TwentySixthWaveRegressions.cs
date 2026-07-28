using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesPendingFormulaFromAnotherSheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellFormula(1, 1, "Data!A5");
            Assert.True(document.HasPendingDirectCellValues);
            ExcelSheet data = AddWorksheetWithoutMaterializingPending(document, "Data");

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
            Assert.True(document.HasPendingDirectCellValues);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanPreflightsPendingFormulaFromAnotherSheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet summary = document.AddWorksheet("Summary");
            for (int column = 1; column <= 128; column++) {
                summary.CellValue(1, column, column);
            }
            summary.CellFormula(1, 129, $"Data!A{A1.MaxRows}");
            Assert.True(document.HasPendingDirectCellValues);
            ExcelSheet data = AddWorksheetWithoutMaterializingPending(document, "Data");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => data.PlanInsertRows(1));

            Assert.Contains("row limit", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(document.HasPendingDirectCellValues);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesNamedSheetViewFilters() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddNamedSheetViewFilter(sheet, "A5:C10");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact views = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "named-sheet-views");
            Assert.Equal(1, views.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesCrossSheetDataConsolidation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.WorksheetPart.Worksheet.Append(
                new DataConsolidate(
                    new DataReferences(
                        new DataReference {
                            Sheet = "Data",
                            Reference = "A5:A6"
                        })) {
                    Function = DataConsolidateFunctionValues.Sum
                });

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact consolidation = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "data-consolidation");
            Assert.Equal(1, consolidation.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanUsesAnchoredTableFormulaSemantics() {
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
            var calculatedFormula = new CalculatedColumnFormula("A3+SUM(A3:A4)+$A$3");
            var totalsFormula = new TotalsRowFormula("$A$3");
            resultColumn.Append(calculatedFormula, totalsFormula);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, formulas.ItemCount);
            Assert.Equal("A3+SUM(A3:A4)+$A$3", calculatedFormula.Text);
            Assert.Equal("$A$3", totalsFormula.Text);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanPrefersTargetSemanticsForSharedCharts() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet firstOwner = CreateChartOwner(document, "First owner");
            ExcelSheet data = CreateChartOwner(document, "Data");
            ChartPart sharedChartPart = Assert.Single(firstOwner.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = sharedChartPart.ChartSpace.Descendants<C.Formula>().First();
            formula.Text = "A5";
            DrawingsPart dataDrawings = data.WorksheetPart.DrawingsPart!;
            ChartPart replacedPart = Assert.Single(dataDrawings.ChartParts);
            string relationshipId = dataDrawings.GetIdOfPart(replacedPart);
            dataDrawings.DeletePart(replacedPart);
            dataDrawings.AddPart(sharedChartPart, relationshipId);

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount == 1);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "drawings" && impact.ItemCount > 0);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesEmptyRowRecords() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheetData.Append(new Row {
                RowIndex = 5U,
                Height = 24D,
                CustomHeight = true
            });

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact rows = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "worksheet-rows");
            Assert.Equal(1, rows.ItemCount);
            Assert.DoesNotContain(plan.Impacts, impact => impact.Category == "worksheet-cells");
        }

        private static ExcelSheet AddWorksheetWithoutMaterializingPending(
            ExcelDocument document,
            string name) {
            WorkbookPart workbookPart = document.OpenXmlDocument.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
            uint sheetId = sheets.Elements<Sheet>().Max(sheet => sheet.SheetId!.Value) + 1U;
            var sheetElement = new Sheet {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            };
            sheets.Append(sheetElement);
            return new ExcelSheet(document, document.OpenXmlDocument, sheetElement);
        }
    }
}
