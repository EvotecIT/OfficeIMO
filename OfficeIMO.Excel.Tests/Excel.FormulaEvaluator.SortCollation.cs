using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("TextAscending", "SORT(A1:B3,1,1)")]
        [InlineData("TextDescending", "SORT(A1:B3,1,-1)")]
        [InlineData("BlankNumeric", "SORT(A1:B3,1,1)")]
        [InlineData("BlankNumericDescending", "SORT(A1:B3,1,-1)")]
        [InlineData("MixedKeys", "SORT(A1:B4,1,1)")]
        [InlineData("MixedDescending", "SORT(A1:B4,1,-1)")]
        [InlineData("MixedBoolean", "SORT(A1:B5,1,1)")]
        [InlineData("MixedBooleanDescending", "SORT(A1:B5,1,-1)")]
        [InlineData("BlankText", "SORT(A1:B3,1,1)")]
        [InlineData("BlankTextDescending", "SORT(A1:B3,1,-1)")]
        [InlineData("TextByColumns", "SORT(A1:C2,1,1,TRUE)")]
        [InlineData("MixedByColumnsDescending", "SORT(A1:E2,1,-1,TRUE)")]
        [InlineData("TextCaseTies", "SORT(A1:B3,1,1)")]
        [InlineData("DigitText", "SORT(A1:B5,1,1)")]
        [InlineData("NumericText", "SORT(A1:B5,1,1)")]
        public void ArraySort_MatchesExcelProducedCachesAfterNewCalculationAndReopen(string caseName, string formula) {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelFormulaCorpus", "sort-collation.xlsx");
            object?[,] input, expected;
            using (var excel = ExcelDocumentReader.Open(source)) {
                var sheet = excel.GetSheet(caseName);
                input = sheet.ReadRange("A1:E5");
                expected = sheet.ReadRange("G1:K5");
            }

            string path = Path.Combine(_directoryWithFiles, "SortCollation-" + caseName + ".xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet(caseName);
                for (int row = 0; row < input.GetLength(0); row++) {
                    for (int column = 0; column < input.GetLength(1); column++) {
                        switch (input[row, column]) {
                            case double number: sheet.CellValue(row + 1, column + 1, number); break;
                            case bool boolean: sheet.CellValue(row + 1, column + 1, boolean); break;
                            case string text: sheet.CellValue(row + 1, column + 1, text); break;
                            case null: break;
                            default: throw new InvalidOperationException("Unexpected Excel oracle input type.");
                        }
                    }
                }
                sheet.SetDynamicArrayFormula("G1", formula);
                Assert.Equal(1, document.Calculate());
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using var actualDocument = ExcelDocumentReader.Open(path);
            object?[,] actual = actualDocument.GetSheet(caseName).ReadRange("G1:K5");
            for (int row = 0; row < expected.GetLength(0); row++) {
                for (int column = 0; column < expected.GetLength(1); column++) {
                    Assert.True(Equals(expected[row, column], actual[row, column]),
                        $"{caseName} {(char)('G' + column)}{row + 1}: Excel={expected[row, column] ?? "<blank>"}; OfficeIMO={actual[row, column] ?? "<blank>"}");
                }
            }
        }
    }
}
