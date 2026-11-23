using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // The vector-form example in Microsoft's LOOKUP reference.
        [Theory]
        [InlineData("LOOKUP(4.19,A2:A6,B2:B6)", "orange")]
        [InlineData("LOOKUP(5.75,A2:A6,B2:B6)", "yellow")]
        [InlineData("LOOKUP(7.66,A2:A6,B2:B6)", "blue")]
        [InlineData("LOOKUP(0,A2:A6,B2:B6)", "#N/A")]
        [InlineData("LOOKUP(5.17,A2:A6)", "5.17")]
        [InlineData("LOOKUP(\"C\",{\"a\",\"b\",\"c\",\"d\";1,2,3,4})", "3")]
        [InlineData("LOOKUP(\"bump\",{\"a\",1;\"b\",2;\"c\",3})", "2")]
        [InlineData("LOOKUP(5,A2:B6)", "orange")]
        public void Test_FormulaEvaluator_LookupMatchesDocumentedExamples(string formula, string expected) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Lookup");
            double[] frequencies = { 4.14, 4.19, 5.17, 5.77, 6.39 };
            string[] colors = { "red", "orange", "yellow", "green", "blue" };
            for (int index = 0; index < frequencies.Length; index++) {
                sheet.Cell(index + 2, 1).SetValue(frequencies[index]);
                sheet.Cell(index + 2, 2).SetValue(colors[index]);
            }
            sheet.Cell(1, 4).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 4, out string? value));
            Assert.Equal(expected, value);
        }

        // The examples in Microsoft's ADDRESS reference.
        [Theory]
        [InlineData("ADDRESS(2,3)", "$C$2")]
        [InlineData("ADDRESS(2,3,2)", "C$2")]
        [InlineData("ADDRESS(2,3,3)", "$C2")]
        [InlineData("ADDRESS(2,3,4)", "C2")]
        [InlineData("ADDRESS(2,3,2,FALSE)", "R2C[3]")]
        [InlineData("ADDRESS(2,3,1,FALSE,\"[Book1]Sheet1\")", "'[Book1]Sheet1'!R2C3")]
        [InlineData("ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")", "'EXCEL SHEET'!R2C3")]
        [InlineData("ADDRESS(1,28,1,TRUE,\"Data\")", "Data!$AB$1")]
        [InlineData("ADDRESS(0,1)", "#VALUE!")]
        [InlineData("ADDRESS(1,1,5)", "#VALUE!")]
        public void Test_FormulaEvaluator_AddressMatchesDocumentedExamples(string formula, string expected) {
            Assert.Equal(expected, CalculateText(formula));
        }

        // The example in Microsoft's INDIRECT reference, with "George" a name for B4.
        [Theory]
        [InlineData("INDIRECT(A2)", "1.333")]
        [InlineData("INDIRECT(A3)", "45")]
        [InlineData("INDIRECT(A4)", "10")]
        [InlineData("INDIRECT(\"B\"&A5)", "62")]
        [InlineData("SUM(INDIRECT(\"B2:B5\"))", "118.333")]
        [InlineData("INDIRECT(\"R3C2\",FALSE)", "45")]
        [InlineData("INDIRECT(\"R[2]C[1]\",FALSE)", "45")]
        [InlineData("INDIRECT(\"Other!A1\")", "7")]
        [InlineData("INDIRECT(\"'Data (2)'!A1\")", "8")]
        [InlineData("INDIRECT(ADDRESS(1,1,1,TRUE,\"Data (2)\"))", "8")]
        [InlineData("INDIRECT(\"George\",FALSE)", "10")]
        [InlineData("INDIRECT(\"B2\",FALSE)", "#REF!")]
        [InlineData("INDIRECT(\"'Missing sheet'!A1\")", "#REF!")]
        [InlineData("INDIRECT(\"not a reference\")", "#REF!")]
        [InlineData("IFERROR(INDIRECT(A5),\"none\")", "none")]
        [InlineData("ROWS(INDIRECT(\"B2:B5\"))", "4")]
        public void Test_FormulaEvaluator_IndirectResolvesReferenceText(string formula, string expected) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            ExcelSheet other = document.AddSheet("Other");
            other.Cell(1, 1).SetValue(7);
            document.AddSheet("Data (2)").Cell(1, 1).SetValue(8);
            sheet.Cell(2, 1).SetValue("B2");
            sheet.Cell(3, 1).SetValue("B3");
            sheet.Cell(4, 1).SetValue("George");
            sheet.Cell(5, 1).SetValue(5);
            sheet.Cell(2, 2).SetValue(1.333);
            sheet.Cell(3, 2).SetValue(45);
            sheet.Cell(4, 2).SetValue(10);
            sheet.Cell(5, 2).SetValue(62);
            document.SetNamedRange("George", "Data!$B$4", save: false);
            sheet.Cell(1, 1).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? value));
            Assert.Equal(expected, value);
        }

        [Fact]
        public void Test_FormulaEvaluator_IndirectFollowsTheCellItNames() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(20);
            sheet.Cell(2, 1).SetFormula("A1*2");
            sheet.Cell(3, 1).SetFormula("INDIRECT(\"A\"&(1+1))+1");
            Assert.Equal(2, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(3, 1, out string? value));
            Assert.Equal("41", value);
        }
    }
}
