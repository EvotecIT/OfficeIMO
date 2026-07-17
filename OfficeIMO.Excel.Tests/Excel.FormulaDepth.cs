using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaDependencyGraph_ReportsDepthEdgesAndCycles() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Graph.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet acyclic = document.AddWorksheet("Acyclic");
                acyclic.CellValue(4, 1, 1d);
                acyclic.CellFormula(3, 1, "A4+1");
                acyclic.CellFormula(2, 1, "A3+1");
                acyclic.CellFormula(1, 1, "A2+1");

                ExcelSheet circular = document.AddWorksheet("Circular");
                circular.CellFormula(1, 1, "B1+1");
                circular.CellFormula(1, 2, "A1+1");
                circular.CellFormula(1, 3, "A1+1");

                ExcelFormulaInspection inspection = document.InspectFormulas();
                ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;

                Assert.Equal(6, graph.NodeCount);
                Assert.Equal(5, graph.EdgeCount);
                Assert.Equal(3, graph.MaximumDependencyDepth);
                Assert.True(graph.HasCircularReferences);

                ExcelFormulaCircularReference cycle = Assert.Single(graph.CircularReferences);
                Assert.Equal(new[] { "Circular!A1", "Circular!B1" }, cycle.References);

                ExcelFormulaDependencyNode acyclicA1 = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Acyclic", "A1"));
                Assert.Equal(3, acyclicA1.DependencyDepth);
                Assert.Equal(new[] { "Acyclic!A2" }, acyclicA1.FormulaDependencies);
                Assert.False(acyclicA1.IsCircular);

                ExcelFormulaDependencyNode circularA1 = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Circular", "A1"));
                Assert.True(circularA1.IsCircular);
                Assert.Null(circularA1.DependencyDepth);

                ExcelFormulaDependencyNode cycleDependent = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Circular", "C1"));
                Assert.False(cycleDependent.IsCircular);
                Assert.Null(cycleDependent.DependencyDepth);

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
                Assert.Contains("Circular!A1 -> Circular!B1", exception.Message, StringComparison.Ordinal);
                Assert.Contains("Maximum dependency depth: 3", graph.ToMarkdown(), StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_StopsAtConfiguredDependencyDepth() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Budget.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Depth");
                sheet.CellValue(4, 1, 1d);
                sheet.CellFormula(3, 1, "A4+1");
                sheet.CellFormula(2, 1, "A3+1");
                sheet.CellFormula(1, 1, "A2+1");
                document.Calculation.MaximumDependencyDepth = 2;

                Assert.Equal(2, document.Calculate());
                Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
                Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? cachedA2));
                Assert.Equal("3", cachedA2);
                Assert.Equal("A2+1", sheet.GetFormulaText(1, 1));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.Sheets[0];
                Assert.Equal("A2+1", sheet.GetFormulaText(1, 1));
                Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaCalculationOptions_ValidateDependencyDepth() {
            var options = new ExcelCalculationOptions();

            Assert.Equal(256, options.MaximumDependencyDepth);
            Assert.Throws<ArgumentOutOfRangeException>(() => options.MaximumDependencyDepth = 0);
            Assert.Equal(256, options.MaximumDependencyDepth);
        }
    }
}
