using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        // MMULT's documented example: {1,3;7,2} × {2,0;0,2}.
        [InlineData("MMULT({1,3;7,2},{2,0;0,2})", "G1:H2", "2|6;14|4")]
        [InlineData("TRANSPOSE({1,2,3;4,5,6})", "G1:H3", "1|4;2|5;3|6")]
        [InlineData("VSTACK({\"A\",\"B\",\"C\"},{\"AA\",\"BB\",\"CC\"})", "G1:I2", "A|B|C;AA|BB|CC")]
        [InlineData("HSTACK({\"A\",\"B\";\"C\",\"D\"},{1,2;3,4})", "G1:J2", "A|B|1|2;C|D|3|4")]
        [InlineData("VSTACK({1,2},{3})", "G1:H2", "1|2;3|#N/A")]
        [InlineData("TEXTSPLIT(\"Dakota Lennon Sanchez\",\" \")", "G1:I1", "Dakota|Lennon|Sanchez")]
        [InlineData("TEXTSPLIT(\"1,2,3;4,5,6\",\",\",\";\")", "G1:I2", "1|2|3;4|5|6")]
        [InlineData("TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,\"-\")", "G1:H2", "a|b;c|-")]
        [InlineData("TEXTSPLIT(\"a,,b\",\",\",,TRUE)", "G1:H1", "a|b")]
        [InlineData("TEXTSPLIT(\"aXbxc\",\"x\",,FALSE,1)", "G1:I1", "a|b|c")]
        [InlineData("TEXTSPLIT(\"a-b.c\",{\"-\",\".\"})", "G1:I1", "a|b|c")]
        [InlineData("SORTBY({\"Tom\";\"Fred\";\"Amy\"},{3;1;2})", "G1:G3", "Fred;Amy;Tom")]
        [InlineData("SORTBY({\"Tom\";\"Fred\";\"Amy\"},{3;1;2},-1)", "G1:G3", "Tom;Amy;Fred")]
        [InlineData("SORTBY({\"a\";\"b\";\"c\";\"d\"},{2;1;2;1},1,{1;2;3;4},-1)", "G1:G4", "d;b;c;a")]
        [InlineData("SORTBY({\"a\",\"b\",\"c\"},{3,1,2})", "G1:I1", "b|c|a")]
        public void Test_ArrayCalculation_ShapingFunctionsFillTheirRange(string formula, string range, string expected) {
            string path = Path.Combine(_directoryWithFiles, "ArrayShaping-" + Guid.NewGuid().ToString("N") + ".xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Arrays");
                sheet.SetArrayFormula(range, formula);
                Assert.Equal(1, document.Calculate());
                document.Save();
            }
            using var reader = ExcelDocumentReader.Open(path);
            object?[,] values = reader.GetSheet("Arrays").ReadRange(range);
            string[][] rows = expected.Split(';').Select(row => row.Split('|')).ToArray();
            Assert.Equal(rows.Length, values.GetLength(0));
            for (int r = 0; r < rows.Length; r++) {
                Assert.Equal(rows[r].Length, values.GetLength(1));
                for (int c = 0; c < rows[r].Length; c++)
                    Assert.Equal(rows[r][c], Convert.ToString(values[r, c], System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        [Theory]
        [InlineData("SUM(MMULT({1,2},{3;4}))", "11")]
        [InlineData("SUM(TRANSPOSE(A1:A3))", "6")]
        [InlineData("COUNTA(VSTACK(A1:A3,B1:B2))", "5")]
        [InlineData("INDEX(HSTACK(A1:A3,B1:B3),2,2)", "20")]
        [InlineData("INDEX(SORTBY(B1:B3,A1:A3,-1),1)", "30")]
        [InlineData("COUNTA(TEXTSPLIT(\"a b c d\",\" \"))", "4")]
        [InlineData("INDEX(MMULT({1,3;7,2},{1,\"x\";0,2}),1,1)", "#VALUE!")]
        [InlineData("INDEX(MMULT({1,2},{1,2}),1,1)", "#VALUE!")]
        [InlineData("SUM(SORTBY(A1:A3,{1;2}))", "#VALUE!")]
        public void Test_ArrayCalculation_ShapingFunctionsFeedScalarFormulas(string formula, string expected) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Arrays");
            for (int row = 1; row <= 3; row++) {
                sheet.Cell(row, 1).SetValue(row);
                sheet.Cell(row, 2).SetValue(row * 10);
            }
            sheet.Cell(1, 4).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 4, out string? value));
            Assert.Equal(expected, value);
        }
    }
}
