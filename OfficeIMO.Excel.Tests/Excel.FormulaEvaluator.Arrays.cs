using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ArrayCalculation_MatchesIndependentExcelCachesAndReopens() {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", "bounded-arrays.xlsx");
            string path = Path.Combine(_directoryWithFiles, "BoundedArrays.xlsx");
            File.Copy(source, path);
            var expected = new Dictionary<string, object?[,]>();
            using (var reader = ExcelDocumentReader.Open(source)) {
                for (int i = 1; i <= 20; i++)
                    expected["Case" + i] = reader.GetSheet("Case" + i).ReadRange("G1:J4");
            }
            using (var document = ExcelDocument.Load(path)) {
                Assert.Equal(106, document.Calculate());
                var deferred = document.InspectFormulas().Formulas.Where(formula => !formula.IsSupportedByOfficeIMO).ToArray();
                Assert.Equal(new[] { "Case19!I1" }, deferred.Select(f => f.SheetName + "!" + f.CellReference).OrderBy(value => value).ToArray());
                document.Save();
                Assert.Empty(document.ValidateOpenXml());
            }
            using var reopened = ExcelDocumentReader.Open(path);
            foreach (var item in expected) {
                var actual = reopened.GetSheet(item.Key).ReadRange("G1:J4");
                for (int r = 0; r < 4; r++) for (int c = 0; c < 4; c++) {
                    Assert.True(Equals(item.Value[r, c], actual[r, c]),
                        item.Key + " [" + r + "," + c + "] Excel=" + item.Value[r, c] + ", OfficeIMO=" + actual[r, c]);
                }
            }
        }

        [Theory]
        [InlineData("SEQUENCE(2,2)", "4")]
        [InlineData("SORT(SEQUENCE(2,2,4,-1),1)", "3")]
        [InlineData("FILTER(SEQUENCE(3,2),SEQUENCE(3)>1)", "6")]
        [InlineData("UNIQUE(SEQUENCE(2,2))", "4")]
        public void Test_ArrayCalculation_FixedRangeAndDependentTail(string formula, string expectedTail) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            // The dependency sorts before the array owner in the worksheet.
            sheet.CellFormula(1, 1, "H3");
            sheet.SetArrayFormula("G2:H3", formula);
            Assert.Equal(2, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal(expectedTail, cached);
        }

        [Fact]
        public void Test_ArrayCalculation_DeferredShapeDoesNotReplaceCaches() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:H2", "SEQUENCE(3,2)");
            sheet.CellValue(2, 8, "preserve");
            Assert.Equal(0, document.Calculate());
            Assert.False(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
        }

        [Fact]
        public void Test_ArrayCalculation_AnotherFormulaInRangeIsPreserved() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:H2", "SEQUENCE(2,2)");
            sheet.CellFormula(2, 8, "99");
            Assert.Equal(1, document.Calculate());
            var formulas = sheet.InspectFormulas().Formulas;
            Assert.False(formulas.Single(f => f.CellReference == "G1").IsSupportedByOfficeIMO);
            Assert.Equal("99", formulas.Single(f => f.CellReference == "H2").CachedValue);
        }

        [Theory]
        [InlineData("SEQUENCE(100001)")]
        [InlineData("SORT(A1:A100001)")]
        public void Test_ArrayCalculation_CellLimitDefersWithoutOutput(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:G2", formula);
            Assert.Equal(0, document.Calculate());
            Assert.Null(Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }

        [Fact]
        public void Test_ArrayCalculation_CircularTailReferenceUsesDependencyGuard() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:G2", "SEQUENCE(2,1,G2)");
            Assert.Equal(0, document.Calculate());
        }

        [Fact]
        public void Test_ArrayCalculation_AnchorReferenceIsScalarAndDoesNotWriteNeighbors() {
            string path = Path.Combine(_directoryWithFiles, "ArrayAnchorReference.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Arrays");
                sheet.CellFormula(1, 1, "G2");
                sheet.CellValue(1, 2, "neighbor");
                sheet.SetArrayFormula("G2:H3", "SEQUENCE(2,2)");
                Assert.Equal(2, document.Calculate());
                document.Save();
            }
            using var reader = ExcelDocumentReader.Open(path);
            var values = reader.GetSheet("Arrays").ReadRange("A1:B2");
            Assert.Equal(1d, values[0, 0]);
            Assert.Equal("neighbor", values[0, 1]);
            Assert.Null(values[1, 0]);
            Assert.Null(values[1, 1]);
        }

        [Fact]
        public void Test_ArrayCalculation_TextLookingLikeNumberOrErrorRetainsItsType() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, "12");
            sheet.CellValue(2, 1, "#N/A");
            sheet.CellFormula(1, 2, "ISNUMBER(A1)");
            sheet.CellFormula(2, 2, "ISERROR(A2)");
            Assert.Equal(2, document.Calculate());
            Assert.All(sheet.InspectFormulas().Formulas, formula => Assert.Equal("0", formula.CachedValue));
        }

        [Fact]
        public void Test_ArrayCalculation_UnsupportedFormulaTextCacheIsNotNumericSortKey() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellFormula(1, 1, "UNSUPPORTED()");
            sheet.CellFormula(2, 1, "UNSUPPORTED()");
            var cells = document.WorkbookPartRoot.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToArray();
            foreach (var cell in cells) {
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(cell.CellReference == "A1" ? "12" : "2");
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }
            sheet.SetArrayFormula("G1:G2", "SORT(A1:A2)");
            Assert.Equal(0, document.Calculate());
        }

        [Fact]
        public void Test_ArrayCalculation_UnsupportedFormulaErrorLookingTextReopensAsText() {
            string path = Path.Combine(_directoryWithFiles, "ArrayTextCache.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Arrays");
                sheet.CellFormula(1, 1, "UNSUPPORTED()");
                var cell = document.WorkbookPartRoot.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("#N/A");
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                sheet.SetArrayFormula("G1:G1", "FILTER(A1:A1,TRUE)");
                Assert.Equal(1, document.Calculate());
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            var output = reopened.WorkbookPartRoot.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(cell => cell.CellReference == "G1");
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.String, output.DataType!.Value);
            Assert.Equal("#N/A", output.CellValue!.Text);
        }

        [Fact]
        public void Test_ArrayCalculation_SmallNonzeroComparisonIsNotEqualToZero() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, 5e-8);
            sheet.CellValue(2, 1, -5e-8);
            sheet.SetArrayFormula("G1:G1", "FILTER(A1:A2,A1:A2>0,\"empty\")");
            sheet.CellFormula(1, 8, "A1=0");
            Assert.Equal(2, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 7, out string? cached));
            Assert.Equal(5e-8, double.Parse(cached!, System.Globalization.CultureInfo.InvariantCulture));
            Assert.True(sheet.TryGetCachedFormulaValue(1, 8, out string? comparison));
            Assert.Equal("0", comparison);
        }

        [Fact]
        public void Test_ArrayCalculation_DeepDependencyChainFitsSmallCallerStackSafely() {
            Exception? failure = null;
            var thread = new System.Threading.Thread(() => {
                try {
                    using var document = ExcelDocument.Create();
                    document.Calculation.MaximumDependencyDepth = 320;
                    var sheet = document.AddWorksheet("Arrays");
                    for (int i = 0; i < 150; i++) {
                        int row = 1 + i * 2;
                        string start = i == 149 ? "1" : "G" + (row + 2) + "+1";
                        sheet.SetArrayFormula("G" + row + ":G" + (row + 1), "SEQUENCE(2,1," + start + ")");
                    }
                    sheet.CellFormula(1, 8, "42");
                    Assert.InRange(document.Calculate(), 1, 151);
                    Assert.True(sheet.TryGetCachedFormulaValue(1, 8, out string? cached));
                    Assert.Equal("42", cached);
                } catch (Exception exception) { failure = exception; }
            }, 1024 * 1024);
            thread.Start();
            Assert.True(thread.Join(TimeSpan.FromSeconds(30)), "Array dependency calculation did not finish.");
            Assert.Null(failure);
        }

        [Theory]
        [InlineData("SEQUENCE(1048577)")]
        [InlineData("SEQUENCE(1,16385)")]
        public void Test_ArrayCalculation_OutsideWorksheetExtentStaysDeferred(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:G1", formula);
            Assert.Equal(0, document.Calculate());
        }

        [Fact]
        public void Test_ArrayCalculation_AtCellBudgetPublishesFinalCell() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:G100000", "SEQUENCE(100000)");
            Assert.Equal(1, document.Calculate());
            Assert.Equal(100000d, sheet.CellAt(100000, 7).GetValue().Value);
        }

        [Fact]
        public void Test_ArrayCalculation_UnqualifiedMixedComparisonStaysDeferred() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, 1d);
            sheet.SetArrayFormula("H1:H1", "FILTER(A1:A1,A1:A1=\"1\",\"none\")");
            Assert.Equal(0, document.Calculate());
        }

        [Fact]
        public void Test_ArrayCalculation_BlankMaskCellIsFalse() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, 2d);
            sheet.CellValue(2, 1, 3d);
            sheet.CellValue(1, 4, true);
            sheet.SetArrayFormula("G1:G1", "FILTER(A1:A2,D1:D2)");
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 7, out string? cached));
            Assert.Equal("2", cached);
        }

        [Fact]
        public void Test_ArrayCalculation_MergedOutputRangeStaysDeferred() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.SetArrayFormula("G1:H2", "SEQUENCE(2,2)");
            document.WorkbookPartRoot.WorksheetParts.Single().Worksheet.AppendChild(
                new DocumentFormat.OpenXml.Spreadsheet.MergeCells(
                    new DocumentFormat.OpenXml.Spreadsheet.MergeCell { Reference = "G1:H1" }));
            Assert.Equal(0, document.Calculate());
        }

        [Fact]
        public void Test_ArrayCalculation_FunctionCatalogReservesArrayNamesWithoutScalarSpilling() {
            using var document = ExcelDocument.Create();
            foreach (string function in new[] { "SEQUENCE", "FILTER", "SORT", "UNIQUE" }) {
                Assert.Contains(function, ExcelFormulaCapabilities.Current.SupportedFunctions);
                Assert.Throws<ArgumentException>(() => document.Calculation.RegisterCustomFunction(function,
                    (context, arguments) => ExcelFormulaValue.FromNumber(1)));
            }
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellFormula(1, 1, "SEQUENCE(2)");
            sheet.CellValue(2, 1, "neighbor");
            Assert.Equal(0, document.Calculate());
            Assert.False(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
            Assert.Equal("neighbor", sheet.CellAt(2, 1).GetValue().Value);
        }

        [Fact]
        public void Test_ArrayCalculation_UnqualifiedUnicodeTextSortStaysDeferred() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, "é");
            sheet.CellValue(2, 1, "a");
            sheet.SetArrayFormula("G1:G2", "SORT(A1:A2)");
            Assert.Equal(0, document.Calculate());
        }
    }
}
