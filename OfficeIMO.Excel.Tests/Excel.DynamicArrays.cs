using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void DynamicArrays_AuthoredMetadataOpensAndSpillsInDesktopExcel() {
            string path = Path.Combine(_directoryWithFiles, "AuthoredDynamicArray.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.SetDynamicArrayFormula("G1", "SEQUENCE(2,2)");
                var inspected = Assert.Single(sheet.InspectFormulas().Formulas);
                Assert.True(inspected.Array?.IsDynamic);
                Assert.Equal("G1:G1", inspected.Array.Range);
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.Empty(document.ValidateOpenXml());
            }
            AssertWorkbookOpensViaExcelComWhenAvailable(path,
                "Authored dynamic-array metadata must open and spill in desktop Excel.",
                new Dictionary<string, string> { ["G1"] = "1", ["H1"] = "2", ["G2"] = "3", ["H2"] = "4" });
        }

        [Fact]
        public void DynamicArrays_RecalculateResizesAndRefreshesDependentCells() {
            string path = Path.Combine(_directoryWithFiles, "ResizedDynamicArray.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 2);
                sheet.CellValue(1, 2, 2);
                sheet.CellFormula(1, 5, "G2+H2");
                sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,B1)");
                document.RecalculateSupportedFormulas();
                Assert.True(sheet.TryGetCellText(1, 7, out var anchor));
                Assert.Equal("1", anchor);
                Assert.True(sheet.TryGetCellText(2, 8, out var child));
                Assert.Equal("4", child);
                Assert.True(sheet.TryGetCachedFormulaValue(1, 5, out var dependent));
                Assert.Equal("7", dependent);
                Assert.Equal("G1:H2", Assert.Single(sheet.InspectFormulas().Formulas, f => f.Array?.IsDynamic == true).Array!.Range);
                Assert.Throws<InvalidOperationException>(() => sheet.CellValue(2, 8, 99));
                Assert.Throws<InvalidOperationException>(() => sheet.CellFormula(2, 8, "1+1"));
                Assert.Throws<InvalidOperationException>(() => sheet.CellValues(new[] { (Row: 2, Column: 8, Value: (object)99) }));
                Assert.Throws<InvalidOperationException>(() => sheet.ClearRange("H2:H2", ExcelClearOptions.Values));

                sheet.CellValue(1, 1, 4);
                document.RecalculateSupportedFormulas();
                Assert.True(sheet.TryGetCellText(4, 8, out var grown));
                Assert.Equal("8", grown);
                Assert.Equal("G1:H4", Assert.Single(sheet.InspectFormulas().Formulas, f => f.Array?.IsDynamic == true).Array!.Range);

                sheet.CellValue(1, 1, 1);
                document.RecalculateSupportedFormulas();
                Assert.False(sheet.TryGetCellText(4, 8, out _));
                Assert.False(sheet.TryGetCellText(2, 8, out _));
                Assert.Equal("G1:H1", Assert.Single(sheet.InspectFormulas().Formulas, f => f.Array?.IsDynamic == true).Array!.Range);
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.Empty(document.ValidateOpenXml());
            }
            using (var reopened = ExcelDocument.Load(path)) {
                var sheet = reopened["Data"];
                Assert.True(sheet.TryGetCellText(1, 8, out var child));
                Assert.Equal("2", child);
                Assert.False(sheet.TryGetCellText(2, 8, out _));
            }
        }

        [Fact]
        public void DynamicArrays_BlockedSpillRecoversWhenBlockerClears() {
            string path = Path.Combine(_directoryWithFiles, "BlockedDynamicArray.xlsx");
            using var document = ExcelDocument.Create(path);
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(2, 8, "blocker");
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(2,2)");
            sheet.CellFormula(1, 9, "G1");
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCachedFormulaValue(1, 7, out var error));
            Assert.Equal("#SPILL!", error);
            Assert.True(sheet.TryGetCellText(2, 8, out var blocker));
            Assert.Equal("blocker", blocker);
            Assert.False(sheet.TryGetCellText(1, 8, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(1, 9, out var propagated));
            Assert.Equal("#SPILL!", propagated);
            sheet.ClearRange("H2:H2", ExcelClearOptions.Values);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(2, 8, out var recovered));
            Assert.Equal("4", recovered);
            Assert.Equal("G1:H2", Assert.Single(sheet.InspectFormulas().Formulas, f => f.Array?.IsDynamic == true).Array!.Range);
            document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void DynamicArrays_ExcelProducedFixtureRetainsCalculatedCases() {
            string source = Path.Combine(_directoryDocuments, "ExcelFormulaCorpus", "dynamic-spills.xlsx");
            string path = Path.Combine(_directoryWithFiles, "CalculatedExcelDynamicSpills.xlsx");
            File.Copy(source, path, overwrite: true);
            using (var document = ExcelDocument.Load(path)) {
                document.Calculate();
                foreach (var (name, address, expected) in new[] {
                    ("Initial", "G1", "1"), ("Grown", "H4", "8"),
                    ("Shrunk", "G1", "1"), ("Blocked", "G1", "#SPILL!"),
                    ("Recovered", "H2", "4"), ("Merged", "G1", "#SPILL!"),
                    ("Table", "G1", "#SPILL!"), ("Edge", "XFD1048576", "#SPILL!"),
                    ("EmptyFilter", "G1", "#CALC!"), ("EmptyUnique", "G1", "#CALC!"),
                    ("ZeroRows", "G1", "#CALC!"), ("ZeroColumns", "G1", "#CALC!"),
                    ("NegativeRows", "G1", "#VALUE!"), ("NegativeColumns", "G1", "#VALUE!")
                }) {
                    var (row, column) = A1.ParseCellRef(address);
                    Assert.True(document[name].TryGetCellText(row, column, out var actual), name);
                    Assert.Equal(expected, actual);
                }
                Assert.False(document["Shrunk"].TryGetCellText(2, 8, out _));
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.Empty(document.ValidateOpenXml());
            }
            using (var reopened = ExcelDocument.Load(path)) {
                Assert.True(reopened["Blocked"].TryGetCachedFormulaValue(1, 7, out var blocked));
                Assert.Equal("#SPILL!", blocked);
                Assert.True(reopened["Initial"].TryGetCellText(2, 8, out var child));
                Assert.Equal("4", child);
            }
        }

        [Fact]
        public void DynamicArrays_OneFreshSpillCanFeedAnotherBeforeEitherIsCached() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(K2,2)");
            sheet.SetDynamicArrayFormula("K1", "SEQUENCE(2,2)");
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(3, 8, out var downstream));
            Assert.Equal("6", downstream);
            Assert.True(sheet.TryGetCellText(2, 12, out var upstream));
            Assert.Equal("4", upstream);
        }

        [Fact]
        public void DynamicArrays_BlockedNativeErrorAndClearOwnership() {
            string path = Path.Combine(_directoryWithFiles, "BlockedNativeDynamicArray.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(2, 8, "blocker");
                sheet.SetDynamicArrayFormula("G1", "SEQUENCE(2,2)");
                sheet.CellFormula(1, 9, "G1");
                document.RecalculateSupportedFormulas();
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.Empty(document.ValidateOpenXml());
            }
            AssertWorkbookOpensViaExcelComWhenAvailable(path,
                "Authored blocked dynamic arrays must remain native Excel spill errors.",
                new Dictionary<string, string> { ["G1"] = "#SPILL!", ["H2"] = "blocker", ["I1"] = "#SPILL!" });
            using (var reopened = ExcelDocument.Load(path)) {
                var sheet = reopened["Data"];
                sheet.ClearArrayFormula("G1");
                sheet.CellValue(1, 7, 7);
                Assert.True(sheet.TryGetCellText(1, 7, out var text));
                Assert.Equal("7", text);
                Assert.DoesNotContain(sheet.InspectFormulas().Formulas, f => f.Array?.IsDynamic == true);
                reopened.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            }
        }

        [Fact]
        public void DynamicArrays_GrowthCollisionRetiresOldChildrenAndRecovers() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 2);
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,2)");
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(2, 8, out var oldChild));
            Assert.Equal("4", oldChild);
            sheet.CellValue(3, 8, "blocker");
            sheet.CellValue(1, 1, 4);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCachedFormulaValue(1, 7, out var error));
            Assert.Equal("#SPILL!", error);
            Assert.False(sheet.TryGetCellText(2, 8, out _));
            Assert.True(sheet.TryGetCellText(3, 8, out var blocker));
            Assert.Equal("blocker", blocker);
            sheet.ClearRange("H3:H3", ExcelClearOptions.Values);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(4, 8, out var recovered));
            Assert.Equal("8", recovered);
        }

        [Fact]
        public void DynamicArrays_HundredThousandCellLimitRemainsBounded() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.SetDynamicArrayFormula("A1", "SEQUENCE(100000,1)");
            Assert.Equal(1, document.RecalculateSupportedFormulas());
            Assert.True(sheet.TryGetCellText(100000, 1, out var final));
            Assert.Equal("100000", final);
            Assert.Equal("A1:A100000", Assert.Single(sheet.InspectFormulas().Formulas).Array!.Range);
            sheet.SetDynamicArrayFormula("C1", "SEQUENCE(100001,1)");
            Assert.Equal(1, document.RecalculateSupportedFormulas());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 3, out _));
            Assert.False(sheet.TryGetCellText(2, 3, out _));
        }
    }
}
