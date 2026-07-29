using System;
using System.Data;
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
        public void Test_StructuralRows_MutationPlanClassifiesExternalValidationAndFormattingFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");

            var validation = new DataValidation(new Formula1("Data!A5")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            var formatting = new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("Data!A5>0")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1
                }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1" }
            };
            summary.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U },
                formatting);

            var extendedValidation = new X14.DataValidation(
                new X14.DataValidationForumla1(new Xm.Formula("Data!A5")),
                new Xm.ReferenceSequence("C1"));
            var extendedFormatting = new X14.ConditionalFormatting(
                new X14.ConditionalFormattingRule(new Xm.Formula("Data!A5>0")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1,
                    Id = "{91D89B06-E47D-42C6-9088-E873D84008F2}"
                },
                new Xm.ReferenceSequence("D1"));
            summary.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(
                        new X14.DataValidations(extendedValidation) { Count = 1U }) {
                        Uri = "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"
                    },
                    new Extension(
                        new X14.ConditionalFormattings(extendedFormatting)) {
                        Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    }));

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact validations = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "validation");
            ExcelMutationImpact formattings = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "conditional-formatting");
            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, validations.ItemCount);
            Assert.Equal(2, formattings.ItemCount);
            Assert.True(formulas.ItemCount >= 4);
            Assert.Equal("Data!A5", validation.Formula1!.Text);
            Assert.Equal("Data!A5>0", formatting.Descendants<Formula>().Single().Text);
            Assert.Equal("Data!A5", extendedValidation.DataValidationForumla1!.Formula!.Text);
            Assert.Equal(
                "Data!A5>0",
                extendedFormatting.Descendants<Xm.Formula>().Single().Text);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsPreservedFastSaveRowsWithoutMaterializing() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.InsertDataTable(CreateDataTable());
            document.PreserveDeferredDataSetFastSaveModelAndClearCandidate();
            Assert.False(document.HasDeferredDirectDataSetImport);
            Assert.True(document.HasUnmaterializedDirectDataSetRows);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanInsertRows(2));

            Assert.Contains("preserved fast-save", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(document.HasUnmaterializedDirectDataSetRows);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanApplyMaterializesPreservedFastSaveRows() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);
            sheet.InsertDataTable(CreateDataTable());
            document.PreserveDeferredDataSetFastSaveModelAndClearCandidate();
            Assert.True(document.HasUnmaterializedDirectDataSetRows);

            plan.Apply();

            Assert.False(document.HasUnmaterializedDirectDataSetRows);
            Assert.True(plan.IsApplied);
            Assert.Equal("Name", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("North", sheet.CellAt(3, 1).GetValue<string>());
        }

        private static DataTable CreateDataTable() {
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("North", 10);
            table.Rows.Add("South", 20);
            return table;
        }
    }
}
