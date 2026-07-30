using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanCountsFormulaRemovedWithDeletedRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(5, 1, "1+1");
            CellFormula formula = Assert.Single(
                sheet.WorksheetPart.Worksheet.Descendants<CellFormula>());
            formula.CalculateCell = true;

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(5);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanCountsFormulaReplacedByPendingOrdinaryValue() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet summary = document.AddWorksheet("Summary");
            SheetData sheetData = summary.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            for (int column = 1; column < 128; column++) {
                summary.CellValue(1, column, column);
            }
            Row row = Assert.Single(sheetData.Elements<Row>());
            row.Append(new Cell {
                CellReference = "DX1",
                CellFormula = new CellFormula("1+1")
            });
            summary.CellValue(1, 128, 128);
            Assert.True(document.HasPendingDirectCellValues);
            Assert.Single(summary.WorksheetPart.Worksheet.Descendants<CellFormula>());
            ExcelSheet data = AddWorksheetWithoutMaterializingPending(document, "Data");

            ExcelRowMutationPlan plan = data.PlanInsertRows(100);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
            plan.Apply();
            CellFormula formula = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<CellFormula>());
            Assert.True(formula.CalculateCell?.Value);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanCountsFormulasRemovedWithMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var validation = new DataValidation(new Formula1("B1")) {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A5"
                }
            };
            var formatting = new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("C1")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1
                }) {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A5"
                }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U },
                formatting);

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(5);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, formulas.ItemCount);
            plan.Apply();
            Assert.Empty(
                sheet.WorksheetPart.Worksheet.Descendants<DataValidation>());
            Assert.Empty(
                sheet.WorksheetPart.Worksheet.Descendants<ConditionalFormatting>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanClassifiesSparklineGroupDateAxisFormula() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(5, 1, 1);
            data.CellValue(6, 1, 2);
            summary.CellValue(1, 1, 1);
            summary.CellValue(2, 1, 2);
            summary.AddSparklines("'Data'!A5:A6", "B1");
            X14.SparklineGroup group = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            X14.Sparkline sparkline = Assert.Single(group.Descendants<X14.Sparkline>());
            sparkline.Formula!.Text = "'Summary'!A1:A2";
            group.Formula = new Xm.Formula("'Data'!A5:A6");

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact sparklines = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "sparklines");
            Assert.Equal(1, sparklines.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanClassifiesTargetSheetSparklineGroupDateAxisFormula() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(5, 1, 5);
            sheet.CellValue(6, 1, 6);
            sheet.AddSparklines("A1:A2", "B1");
            X14.SparklineGroup group = Assert.Single(
                sheet.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            group.Formula = new Xm.Formula("A5:A6");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact sparklines = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "sparklines");
            Assert.Equal(1, sparklines.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsRemovedWorksheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet original = document.AddWorksheet("Data");
            ExcelRowMutationPlan plan = original.PlanInsertRows(1);
            document.RemoveWorksheet("Data");
            ExcelSheet replacement = document.AddWorksheet("Data");
            replacement.CellValue(1, 1, "Replacement");

            InvalidOperationException exception =
                Assert.Throws<InvalidOperationException>(() => plan.Apply());

            Assert.Contains("no longer part", exception.Message);
            Assert.Equal("Replacement", replacement.CellAt(1, 1).GetValue<string>());
            Assert.True(plan.IsConsumed);
            Assert.False(plan.IsApplied);
        }

        [Fact]
        public void Test_StructuralRows_UnchangedMergeContainerIsNotNormalized() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var merges = new MergeCells(new MergeCell { Reference = "A1:B1" });
            sheet.WorksheetPart.Worksheet.Append(merges);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "merged-cells");
            plan.Apply();
            Assert.Null(merges.Count);
            Assert.Equal(
                "A1:B1",
                Assert.Single(merges.Elements<MergeCell>()).Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanChargesNamedPivotLookup() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            WorksheetSource source = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            source.Reference = null;
            source.Sheet = null;
            source.Name = "MissingPivotSource";
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                new DefinedName("'Data'!$A$1:$A$2") { Name = "OtherName" });

            ExcelRowMutationPlan namedPlan = sheet.PlanInsertRows(5);
            source.Name = null;
            ExcelRowMutationPlan unnamedPlan = sheet.PlanInsertRows(5);

            Assert.True(namedPlan.ScannedElements > unnamedPlan.ScannedElements);
        }
    }
}
