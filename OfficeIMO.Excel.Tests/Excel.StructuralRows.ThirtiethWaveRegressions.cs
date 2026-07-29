using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresUnchangedValidationAndFormattingRules() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var validation = new DataValidation {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A1:A2"
                }
            };
            var formatting = new ConditionalFormatting(
                new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.DuplicateValues,
                    Priority = 1
                }) {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "B1:B2"
                }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U },
                formatting);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(100);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "validation");
            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "conditional-formatting");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesUnchangedSharedFormulaMaterialization() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet other = document.AddWorksheet("Other");
            other.CellAt(2, 1).SetValue(1);
            other.CellAt(3, 1).SetValue(2);
            AppendSharedFormulaGroup(
                other,
                sharedIndex: 71U,
                anchorReference: "B2:B3");
            Cell anchor = other.WorksheetPart.Worksheet
                .Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B2");
            anchor.CellFormula!.Text = "1+1";

            ExcelRowMutationPlan plan = data.PlanInsertRows(100);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, formulas.ItemCount);
        }
    }
}
