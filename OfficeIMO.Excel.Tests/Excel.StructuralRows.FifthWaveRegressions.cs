using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MapsAnchoredFormulaTargetsIndependentlyFromRuleAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(3);

            var stationaryRule = new DataValidation(new Formula1("A3>0")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1" }
            };
            var movingRule = new DataValidation(new Formula1("A1>0")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B3" }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(stationaryRule, movingRule) { Count = 2U });

            sheet.InsertRows(2);

            Assert.Equal("B1", stationaryRule.SequenceOfReferences!.InnerText);
            Assert.Equal("A4>0", stationaryRule.Formula1!.Text);
            Assert.Equal("B4", movingRule.SequenceOfReferences!.InnerText);
            Assert.Equal("A1>0", movingRule.Formula1!.Text);

            sheet.DeleteRows(2);

            Assert.Equal("B1", stationaryRule.SequenceOfReferences!.InnerText);
            Assert.Equal("A3>0", stationaryRule.Formula1!.Text);
            Assert.Equal("B3", movingRule.SequenceOfReferences!.InnerText);
            Assert.Equal("A1>0", movingRule.Formula1!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RewritesUnicodeBareSheetQualifiers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("数据");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(3, 1).SetValue(3);
            summary.CellFormula(1, 1, "数据!A3");

            data.InsertRows(3);

            Assert.Equal("数据!A4", summary.GetFormulaText(1, 1));
        }

        [Fact]
        public void Test_StructuralRows_RemapsWorkbookDataConsolidationSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue(5);
            data.CellAt(6, 1).SetValue(6);

            var internalSource = new DataReference {
                Sheet = "Data",
                Reference = "A5:A6"
            };
            var relationshipSource = new DataReference {
                Sheet = "Data",
                Reference = "B5:B6",
                Id = "rIdExternal"
            };
            var sources = new DataReferences(internalSource, relationshipSource) { Count = 2U };
            summary.WorksheetPart.Worksheet.Append(
                new DataConsolidate(sources) {
                    Function = DataConsolidateFunctionValues.Sum
                });

            data.InsertRows(5);

            Assert.Equal("A6:A7", internalSource.Reference!.Value);
            Assert.Equal("B5:B6", relationshipSource.Reference!.Value);
            Assert.Equal(2U, sources.Count!.Value);

            data.DeleteRows(5);

            Assert.Equal("A5:A6", internalSource.Reference!.Value);
            Assert.Equal("B5:B6", relationshipSource.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_ShiftsAndRemovesCellWatches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(5);
            var watch = new CellWatch { CellReference = "A5" };
            sheet.WorksheetPart.Worksheet.Append(new CellWatches(watch));

            sheet.InsertRows(5);

            Assert.Equal("A6", watch.CellReference!.Value);

            sheet.DeleteRows(6);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<CellWatches>());
        }
    }
}
