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

                ExcelSheet dependenciesFirst = document.AddWorksheet("DependenciesFirst");
                dependenciesFirst.CellValue(1, 1, 1d);
                dependenciesFirst.CellFormula(2, 1, "A1+1");
                dependenciesFirst.CellFormula(3, 1, "A2+1");
                dependenciesFirst.CellFormula(4, 1, "A3+1");
                document.Calculation.MaximumDependencyDepth = 2;

                Assert.Equal(4, document.Calculate());
                Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
                Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? cachedA2));
                Assert.Equal("3", cachedA2);
                Assert.Equal("A2+1", sheet.GetFormulaText(1, 1));
                Assert.False(dependenciesFirst.TryGetCachedFormulaValue(4, 1, out _));
                Assert.True(dependenciesFirst.TryGetCachedFormulaValue(3, 1, out string? cachedForwardA3));
                Assert.Equal("3", cachedForwardA3);
                Assert.Equal("A3+1", dependenciesFirst.GetFormulaText(4, 1));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.Sheets[0];
                Assert.Equal("A2+1", sheet.GetFormulaText(1, 1));
                Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
                ExcelSheet dependenciesFirst = document.Sheets[1];
                Assert.Equal("A3+1", dependenciesFirst.GetFormulaText(4, 1));
                Assert.False(dependenciesFirst.TryGetCachedFormulaValue(4, 1, out _));
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesDefinedNamesAndStructuredReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Aliases.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet named = document.AddWorksheet("Named Circular");
                named.CellFormula(1, 1, "Loop+1");
                document.SetNamedRange("Loop", "'Named Circular'!A1", save: false);

                ExcelSheet structured = document.AddWorksheet("Structured");
                structured.CellValue(1, 1, "Amount");
                structured.CellFormula(2, 1, "SUM(Sales[Amount])");
                structured.AddTable("A1:A2", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);

                ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
                Assert.Equal(2, graph.NodeCount);
                Assert.Equal(2, graph.EdgeCount);
                Assert.Equal(2, graph.CircularReferenceCount);

                ExcelFormulaDependencyNode namedNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Named Circular", "A1"));
                Assert.Equal(new[] { "Named Circular!A1" }, namedNode.Dependencies);
                Assert.Equal(new[] { "Named Circular!A1" }, namedNode.FormulaDependencies);
                Assert.True(namedNode.IsCircular);

                ExcelFormulaDependencyNode structuredNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Structured", "A2"));
                Assert.Equal(new[] { "Structured!A2" }, structuredNode.Dependencies);
                Assert.Equal(new[] { "Structured!A2" }, structuredNode.FormulaDependencies);
                Assert.True(structuredNode.IsCircular);

                Assert.Throws<InvalidOperationException>(() => document.InspectFormulas().EnsureNoDependencyIssues());
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_ClearsPreExistingCacheWhenLowerDepthBlocksFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.PreExistingCache.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Depth");
                sheet.CellValue(1, 1, 1d);
                sheet.CellFormula(2, 1, "A1+1");
                sheet.CellFormula(3, 1, "A2+1");
                sheet.CellFormula(4, 1, "A3+1");

                document.Calculation.MaximumDependencyDepth = 4;
                Assert.Equal(3, document.Calculate());
                Assert.True(sheet.TryGetCachedFormulaValue(4, 1, out string? initialA4));
                Assert.Equal("4", initialA4);

                document.Calculation.MaximumDependencyDepth = 2;
                Assert.Equal(2, document.Calculate());
                Assert.False(sheet.TryGetCachedFormulaValue(4, 1, out _));
                ExcelFormulaCellInfo blocked = Assert.Single(document.InspectFormulas().Formulas, formula => formula.CellReference == "A4");
                Assert.True(blocked.IsDirty);
                Assert.Equal("A3+1", blocked.Formula);
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
