using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_IndependentChainsRespectBudgetAndRecover() {
            // Dependency-budget evidence needs a known sufficient caller stack.
            // Small-stack safety is exercised separately below on every runtime.
            Exception? failure = null;
            var worker = new System.Threading.Thread(() => {
                try {
                    string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", "dependency-conformance.xlsx");
                    string path = Path.Combine(_directoryWithFiles, "DependencyConformance.xlsx");
                    File.Copy(source, path);
                    var producer = new Dictionary<string, string?>();
                    using (var original = ExcelDocument.Load(source)) {
                        foreach (var cell in original.InspectFormulas().Formulas)
                            producer.Add(cell.CellReference, cell.CachedValue);
                        Assert.Equal(600, producer.Count);
                        Assert.Equal("301", producer["A301"]);
                        Assert.Equal("301", producer["B1"]);
                    }
                    using (var document = ExcelDocument.Load(path)) {
                        Assert.Equal(300, document.InspectFormulas().DependencyGraph.MaximumDependencyDepth);
                        Assert.Equal(256, document.Calculation.MaximumDependencyDepth);
                        int boundedCount = document.Calculate();
                        Assert.True(boundedCount == 512, $"Calculated {boundedCount}; cleared: {string.Join(",", document.InspectFormulas().Formulas.Where(cell => cell.CachedValue == null).Select(cell => cell.CellReference))}");
                        foreach (var cell in document.InspectFormulas().Formulas) {
                            int row = int.Parse(cell.CellReference.Substring(1));
                            int depth = cell.CellReference[0] == 'A' ? row - 1 : 301 - row;
                            if (depth > 256) {
                                Assert.Null(cell.CachedValue);
                                Assert.True(cell.IsDirty);
                            } else {
                                Assert.Equal(producer[cell.CellReference], cell.CachedValue);
                            }
                        }
                        // A subsequent higher budget must restore every independently produced cache.
                        document.Calculation.MaximumDependencyDepth = 320;
                        Assert.Equal(600, document.Calculate());
                        foreach (var cell in document.InspectFormulas().Formulas)
                            Assert.Equal(producer[cell.CellReference], cell.CachedValue);
                        document.Save();
                        Assert.Empty(document.ValidateOpenXml());
                    }
                    using var reopened = ExcelDocument.Load(path);
                    foreach (var cell in reopened.InspectFormulas().Formulas)
                        Assert.Equal(producer[cell.CellReference], cell.CachedValue);
                    // Calculation options are an in-memory policy; reload uses the safe default.
                    Assert.Equal(512, reopened.Calculate());
                    Assert.Equal(88, reopened.InspectFormulas().Formulas.Count(cell => cell.CachedValue == null && cell.IsDirty));
                } catch (Exception exception) { failure = exception; }
            }, 8 * 1024 * 1024);
            worker.Start();
            worker.Join();
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }

        [Fact]
        public void Test_FormulaEvaluator_IndependentCycleIsReportedAndDoesNotPoisonOtherCells() {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", "circular-conformance.xlsx");
            string path = Path.Combine(_directoryWithFiles, "CircularConformance.xlsx");
            File.Copy(source, path);
            using (var document = ExcelDocument.Load(path)) {
                var original = document.InspectFormulas().Formulas.ToDictionary(cell => cell.CellReference, cell => cell.Formula);
                var graph = document.InspectFormulas().DependencyGraph;
                Assert.Equal(new[] { "Circular!A1", "Circular!B1" }, Assert.Single(graph.CircularReferences).References);
                Assert.Null(graph.FindNode("Circular", "C1")!.DependencyDepth);
                Assert.False(graph.FindNode("Circular", "C1")!.IsCircular);
                Assert.Throws<InvalidOperationException>(() => document.InspectFormulas().EnsureNoDependencyIssues());
                Assert.Equal(1, document.Calculate());
                foreach (var cell in document.InspectFormulas().Formulas) {
                    Assert.Equal(original[cell.CellReference], cell.Formula);
                    if (cell.CellReference == "D1") Assert.Equal("42", cell.CachedValue);
                    else {
                        Assert.Null(cell.CachedValue);
                        Assert.True(cell.IsDirty);
                    }
                }
                document.Save();
                Assert.Empty(document.ValidateOpenXml());
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Single(reopened.InspectFormulas().DependencyGraph.CircularReferences);
            Assert.Equal(3, reopened.InspectFormulas().Formulas.Count(cell => cell.CachedValue == null && cell.IsDirty));
            Assert.Equal("42", reopened.InspectFormulas().Formulas.Single(cell => cell.CellReference == "D1").CachedValue);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        public void Test_FormulaEvaluator_ComposedDependenciesRespectCallerStack(bool crossSheet, bool parentheses) {
            Exception? failure = null;
            var worker = new System.Threading.Thread(() => {
                try {
                    using var document = ExcelDocument.Create();
                    var first = document.AddWorksheet("First");
                    var second = crossSheet ? document.AddWorksheet("Second") : first;
                    const int cells = 30;
                    int nesting = parentheses ? 70 : 50;
                    for (int row = 1; row <= cells; row++) {
                        var sheet = row % 2 == 0 ? second : first;
                        var next = row % 2 == 0 ? first : second;
                        string child = row == cells ? "1" : next.Name + "!A" + (row + 1);
                        string prefix = parentheses ? new string('(', nesting) : string.Concat(Enumerable.Repeat("ABS(", nesting));
                        sheet.CellFormula(row, 1, prefix + child + new string(')', nesting));
                    }
                    first.CellFormula(1, 2, "40+2");
                    Assert.True(document.Calculate() > 0);
                    Assert.True(first.TryGetCachedFormulaValue(1, 2, out string? unrelated));
                    Assert.Equal("42", unrelated);
                    var guarded = document.InspectFormulas().Formulas.Where(cell => cell.CellReference != "B1").ToArray();
                    Assert.Contains(guarded, cell => cell.CachedValue == null && cell.IsDirty);
                    Assert.Equal(cells, guarded.Length);
                    Assert.All(guarded, cell => Assert.False(string.IsNullOrEmpty(cell.Formula)));
                } catch (Exception exception) { failure = exception; }
            }, 1024 * 1024);
            worker.Start();
            worker.Join();
            Assert.Null(failure);
        }
    }
}
