using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesExternalFormulaThresholds() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            AppendFormulaThreshold(summary, "C2:C3", "Data!A5");

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount == 1);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "conditional-formatting" && impact.ItemCount == 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesProtectedRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var protectedRange = new ProtectedRange {
                Name = "Editable",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A5:A6" }
            };
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertAfter(
                new ProtectedRanges(protectedRange),
                sheetData);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact ranges = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "protected-ranges");
            Assert.Equal(1, ranges.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsNestedCommentXml() {
            int plainScan = PlanWithLegacyCommentRuns(1);
            int richScan = PlanWithLegacyCommentRuns(64);

            Assert.True(
                richScan - plainScan >= 100,
                $"Expected nested comment XML to consume the scan budget; plain={plainScan}, rich={richScan}.");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanUsesChargedConditionalFormattingSnapshot() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var formatting = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A2:A5"
                }
            };
            for (int index = 0; index < 64; index++) {
                formatting.Append(new ConditionalFormattingRule(
                    new Formula($"B{index + 1}")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = index + 1
                });
            }
            sheet.WorksheetPart.Worksheet.Append(formatting);
            ExcelRowMutationPlan baseline = sheet.PlanDeleteRows(2, 2);

            ExcelRowMutationPlan exactBudget = sheet.PlanDeleteRows(
                2,
                2,
                new ExcelMutationPlanOptions {
                    MaximumScannedElements = baseline.ScannedElements
                });
            InvalidOperationException exception =
                Assert.Throws<InvalidOperationException>(() =>
                    sheet.PlanDeleteRows(
                        2,
                        2,
                        new ExcelMutationPlanOptions {
                            MaximumScannedElements = baseline.ScannedElements - 1
                        }));

            Assert.Equal(baseline.ScannedElements, exactBudget.ScannedElements);
            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanUsesAnchoredValidationAndFormattingSemantics() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var validation = new DataValidation(new Formula1("B1")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A5" }
            };
            var formatting = new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("B1")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1
                }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A5" }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U },
                formatting);

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(2, 2);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, formulas.ItemCount);
            Assert.Equal("B1", validation.Formula1!.Text);
            Assert.Equal("B1", formatting.Descendants<Formula>().Single().Text);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesWebPublishRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            WorkbookPart workbookPart = sheet.WorksheetPart.GetParentParts()
                .OfType<WorkbookPart>()
                .Single();
            var item = new WebPublishItem {
                Id = 1U,
                DivId = "Data_1",
                SourceType = WebSourceValues.Range,
                SourceObject = "Data",
                SourceRef = "A5:B6",
                DestinationFile = "published.htm"
            };
            workbookPart.Workbook.Append(new WebPublishItems(item) { Count = 1U });

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact publish = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "web-publish");
            Assert.Equal(1, publish.ItemCount);
        }

        private static int PlanWithLegacyCommentRuns(int runCount) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetCommentRichText(
                2,
                1,
                Enumerable.Range(0, runCount)
                    .Select(index => new ExcelRichTextRun("Run " + index)));
            return sheet.PlanInsertRows(2).ScannedElements;
        }
    }
}
