using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using ExcelTableStyle = OfficeIMO.Excel.ExcelTableStyle;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaDependencyGraph_ReportsDepthEdgesAndCycles() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Graph.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet acyclic = document.AddSheet("Acyclic");
                acyclic.Cell(4, 1).SetValue(1d);
                acyclic.Cell(3, 1).SetFormula("A4+1");
                acyclic.Cell(2, 1).SetFormula("A3+1");
                acyclic.Cell(1, 1).SetFormula("A2+1");

                ExcelSheet circular = document.AddSheet("Circular");
                circular.Cell(1, 1).SetFormula("B1+1");
                circular.Cell(1, 2).SetFormula("A1+1");
                circular.Cell(1, 3).SetFormula("A1+1");

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

                ExcelFeatureFinding dependencyFinding = document.InspectFeatures().Features
                    .Single(feature => feature.Name == "Formula dependency issues");
                Assert.Contains(dependencyFinding.Details, detail =>
                    detail.Contains("Circular reference: Circular!A1 -> Circular!B1", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesWholeRowAndColumnReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet columns = document.AddSheet("Whole Columns");
            columns.Cell(1, 1).SetFormula("SUM(B:B)");
            columns.Cell(1, 2).SetFormula("A1");
            ExcelSheet rows = document.AddSheet("Whole Rows");
            rows.Cell(1, 1).SetFormula("SUM(2:2)");
            rows.Cell(2, 1).SetFormula("A1");

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            Assert.Equal(4, graph.NodeCount);
            Assert.Equal(4, graph.EdgeCount);
            Assert.Equal(2, graph.CircularReferenceCount);

            ExcelFormulaDependencyNode columnNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Whole Columns", "A1"));
            Assert.Equal(new[] { "Whole Columns!B:B" }, columnNode.Dependencies);
            Assert.Equal(new[] { "Whole Columns!B1" }, columnNode.FormulaDependencies);

            ExcelFormulaDependencyNode rowNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Whole Rows", "A1"));
            Assert.Equal(new[] { "Whole Rows!2:2" }, rowNode.Dependencies);
            Assert.Equal(new[] { "Whole Rows!A2" }, rowNode.FormulaDependencies);
        }

        [Fact]
        public void Test_FormulaInspectionRejectsOversizedFormulaBeforeDependencyMasking() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Bounds");
            sheet.Cell(1, 1).SetValue(1D);
            Cell cell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>());
            cell.CellFormula = new CellFormula("SUM(" + new string('A', 9000) + ")");

            ExcelFormulaCellInfo formula = Assert.Single(document.InspectFormulas().Formulas);

            Assert.Empty(formula.Dependencies);
            Assert.Contains("longer than", formula.UnsupportedReason, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_FormulaDependencyGraphCapsWholeColumnExpansionPerFormula() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Dense");
            for (int row = 1; row <= 5000; row++) sheet.Cell(row, 1).SetFormula("1");
            sheet.Cell(1, 2).SetFormula("SUM(A:A)");

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyNode node = Assert.IsType<ExcelFormulaDependencyNode>(
                inspection.DependencyGraph.FindNode("Dense", "B1"));

            Assert.Equal(4096, node.FormulaDependencies.Count);
            Assert.True(inspection.DependencyGraph.AnalysisTruncated);
            Assert.True(inspection.HasDependencyIssues);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_IgnoresReferenceShapeFunctionArguments() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Reference Shapes");
            sheet.Cell(1, 1).SetFormula("ROW(A1)");
            sheet.Cell(1, 2).SetFormula("COLUMN(B1)");
            sheet.Cell(1, 3).SetFormula("ROWS(C1:C3)");
            sheet.Cell(1, 4).SetFormula("COLUMNS(D1:F1)");
            sheet.Cell(1, 5).SetFormula("SUM(ROW(E1),F1)");
            sheet.Cell(1, 6).SetFormula("1+1");
            sheet.Cell(1, 7).SetFormula("ROWS(ShapeRows)");
            sheet.Cell(1, 8).SetFormula("COLUMNS(ShapeTable[[ROW(A1)]:[B]])");
            sheet.Cell(1, 9).SetFormula("SUM(ShapeTable[ROW(A1)])");
            sheet.Cell(1, 10).SetValue("ROW(A1)");
            sheet.Cell(1, 11).SetValue("B");
            sheet.Cell(2, 10).SetFormula("1+1");
            sheet.Cell(2, 11).SetFormula("ROWS([@B])");
            sheet.AddTable("J1:K2", hasHeader: true, name: "ShapeTable", style: ExcelTableStyle.TableStyleMedium2);
            document.SetNamedRange("ShapeRows", "'Reference Shapes'!C1:C3", save: false);

            ExcelSheet quotedShapeName = document.AddSheet("ROW(A1)");
            quotedShapeName.Cell(1, 2).SetFormula("1+1");
            sheet.Cell(1, 12).SetFormula("'ROW(A1)'!B1");
            sheet.Cell(1, 13).SetFormula("ROWS(M:M)");
            sheet.Cell(1, 14).SetFormula("COLUMNS(1:1)");
            sheet.Cell(1, 15).SetFormula("ROWS('Reference Shapes'!O:O)");
            sheet.Cell(1, 16).SetFormula("COLUMNS('Reference Shapes'!1:1)");

            int calculated = document.Calculate();
            Assert.True(sheet.TryGetCachedFormulaValue(1, 13, out string? directWholeColumnRows));
            Assert.Equal(ExcelA1.MaxRows.ToString(), directWholeColumnRows);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 14, out string? directWholeRowColumns));
            Assert.Equal(ExcelA1.MaxColumns.ToString(), directWholeRowColumns);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 15, out string? qualifiedWholeColumnRows));
            Assert.Equal(ExcelA1.MaxRows.ToString(), qualifiedWholeColumnRows);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 16, out string? qualifiedWholeRowColumns));
            Assert.Equal(ExcelA1.MaxColumns.ToString(), qualifiedWholeRowColumns);
            ExcelFormulaInspection inspection = document.InspectFormulas();
            string[] unsupported = inspection.Formulas
                .Where(formula => !formula.IsSupportedByOfficeIMO)
                .Select(formula => $"{formula.SheetName}!{formula.CellReference}: {formula.UnsupportedReason}")
                .ToArray();
            Assert.True(unsupported.Length == 0, string.Join(Environment.NewLine, unsupported));
            string[] missingCaches = inspection.Formulas
                .Where(formula => !formula.HasCachedValue)
                .Select(formula => $"{formula.SheetName}!{formula.CellReference}: {formula.Formula}")
                .ToArray();
            Assert.True(missingCaches.Length == 0, string.Join(Environment.NewLine, missingCaches));
            Assert.Equal(17, calculated);
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            foreach (string reference in new[] { "A1", "B1", "C1", "D1", "G1", "H1", "M1", "N1", "O1", "P1" }) {
                ExcelFormulaDependencyNode node = Assert.IsType<ExcelFormulaDependencyNode>(
                    graph.FindNode("Reference Shapes", reference));
                Assert.Empty(node.Dependencies);
                Assert.Empty(node.FormulaDependencies);
                Assert.False(node.IsCircular);
            }

            ExcelFormulaDependencyNode mixed = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Reference Shapes", "E1"));
            Assert.Equal(new[] { "Reference Shapes!F1" }, mixed.Dependencies);
            Assert.Equal(new[] { "Reference Shapes!F1" }, mixed.FormulaDependencies);
            ExcelFormulaDependencyNode structuredHeader = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Reference Shapes", "I1"));
            Assert.Equal(new[] { "Reference Shapes!J2" }, structuredHeader.Dependencies);
            Assert.Equal(new[] { "Reference Shapes!J2" }, structuredHeader.FormulaDependencies);
            ExcelFormulaDependencyNode unqualifiedShape = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Reference Shapes", "K2"));
            Assert.Empty(unqualifiedShape.Dependencies);
            Assert.Empty(unqualifiedShape.FormulaDependencies);
            Assert.False(unqualifiedShape.IsCircular);
            ExcelFormulaDependencyNode quotedQualifier = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Reference Shapes", "L1"));
            Assert.Equal(new[] { "ROW(A1)!B1" }, quotedQualifier.Dependencies);
            Assert.Equal(new[] { "ROW(A1)!B1" }, quotedQualifier.FormulaDependencies);
            Assert.Equal(3, graph.EdgeCount);
            Assert.False(graph.HasCircularReferences);
            inspection.EnsureNoDependencyIssues();
        }

        [Fact]
        public void Test_FormulaDependencyGraph_RespectsRangeIntersections() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Intersections");
            sheet.Cell(1, 1).SetFormula("D1");
            sheet.Cell(1, 2).SetFormula("D3");
            sheet.Cell(2, 2).SetFormula("1+1");
            sheet.Cell(1, 4).SetFormula("SUM(A1:B2 B2:C3)");
            sheet.Cell(2, 4).SetFormula("SUM(Data B2:C3)");
            sheet.Cell(3, 4).SetFormula("SUM(A1 (A1:B2))");
            sheet.Cell(4, 4).SetFormula("SUM((A1:B2) B2:C3)");
            sheet.Cell(5, 4).SetFormula("SUM((A1:B2) (B2:C3))");
            sheet.Cell(6, 4).SetFormula("SUM((A1:B2 B2:C3) B2:D2)");
            sheet.Cell(2, 5).SetFormula("ROWS(D1:E2 E2:F3)");
            sheet.Cell(3, 5).SetFormula("SUM(ROWS((D1:E2 E2:F3)),D1)");
            document.SetNamedRange("Data", "Intersections!A1:B2", save: false);

            ExcelSheet structuredSheet = document.AddSheet("Structured Intersections");
            structuredSheet.Cell(1, 1).SetValue("Amount");
            structuredSheet.Cell(1, 2).SetValue("Other");
            structuredSheet.Cell(2, 1).SetFormula("1+1");
            structuredSheet.Cell(2, 2).SetValue(3d);
            structuredSheet.Cell(1, 4).SetFormula("SUM(IntersectionData[Amount] A2:B2)");
            structuredSheet.AddTable(
                "A1:B2",
                hasHeader: true,
                name: "IntersectionData",
                style: ExcelTableStyle.TableStyleMedium2);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode intersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D1"));
            Assert.Equal(new[] { "Intersections!B2" }, intersection.Dependencies);
            Assert.Equal(new[] { "Intersections!B2" }, intersection.FormulaDependencies);
            ExcelFormulaDependencyNode namedIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D2"));
            Assert.Equal(new[] { "Intersections!B2" }, namedIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!B2" }, namedIntersection.FormulaDependencies);
            ExcelFormulaDependencyNode parenthesizedIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D3"));
            Assert.Equal(new[] { "Intersections!A1" }, parenthesizedIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!A1" }, parenthesizedIntersection.FormulaDependencies);
            ExcelFormulaDependencyNode leftParenthesizedIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D4"));
            Assert.Equal(new[] { "Intersections!B2" }, leftParenthesizedIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!B2" }, leftParenthesizedIntersection.FormulaDependencies);
            ExcelFormulaDependencyNode bothParenthesizedIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D5"));
            Assert.Equal(new[] { "Intersections!B2" }, bothParenthesizedIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!B2" }, bothParenthesizedIntersection.FormulaDependencies);
            ExcelFormulaDependencyNode compositeParenthesizedIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "D6"));
            Assert.Equal(new[] { "Intersections!B2" }, compositeParenthesizedIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!B2" }, compositeParenthesizedIntersection.FormulaDependencies);
            ExcelFormulaDependencyNode shapeIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "E2"));
            Assert.Empty(shapeIntersection.Dependencies);
            Assert.Empty(shapeIntersection.FormulaDependencies);
            Assert.False(shapeIntersection.IsCircular);
            ExcelFormulaDependencyNode groupedShapeIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Intersections", "E3"));
            Assert.Equal(new[] { "Intersections!D1" }, groupedShapeIntersection.Dependencies);
            Assert.Equal(new[] { "Intersections!D1" }, groupedShapeIntersection.FormulaDependencies);
            Assert.False(groupedShapeIntersection.IsCircular);
            ExcelFormulaDependencyNode structuredIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Intersections", "D1"));
            Assert.Equal(new[] { "Structured Intersections!A2" }, structuredIntersection.Dependencies);
            Assert.Equal(new[] { "Structured Intersections!A2" }, structuredIntersection.FormulaDependencies);
            Assert.Equal(10, graph.EdgeCount);
            Assert.False(graph.HasCircularReferences);
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesCurrentRowStructuredReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Current Row");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetFormula("CurrentRowData[@B]");
            sheet.Cell(2, 2).SetFormula("CurrentRowData[[#This Row],[A]]");
            sheet.AddTable("A1:B2", hasHeader: true, name: "CurrentRowData", style: ExcelTableStyle.TableStyleMedium2);

            ExcelSheet crossSource = document.AddSheet("Cross Source");
            crossSource.Cell(2, 1).SetFormula("CrossRowData[@B]");
            ExcelSheet crossTable = document.AddSheet("Cross Table");
            crossTable.Cell(1, 1).SetValue("A");
            crossTable.Cell(1, 2).SetValue("B");
            crossTable.Cell(2, 1).SetValue(1d);
            crossTable.Cell(2, 2).SetFormula("'Cross Source'!A2");
            crossTable.AddTable("A1:B2", hasHeader: true, name: "CrossRowData", style: ExcelTableStyle.TableStyleMedium2);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode first = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Current Row", "A2"));
            Assert.Equal(new[] { "Current Row!B2" }, first.Dependencies);
            Assert.Equal(new[] { "Current Row!B2" }, first.FormulaDependencies);
            ExcelFormulaDependencyNode second = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Current Row", "B2"));
            Assert.Equal(new[] { "Current Row!A2" }, second.Dependencies);
            Assert.Equal(new[] { "Current Row!A2" }, second.FormulaDependencies);
            ExcelFormulaDependencyNode crossFirst = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Cross Source", "A2"));
            Assert.Equal(new[] { "Cross Table!B2" }, crossFirst.Dependencies);
            Assert.Equal(new[] { "Cross Table!B2" }, crossFirst.FormulaDependencies);
            ExcelFormulaDependencyNode crossSecond = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Cross Table", "B2"));
            Assert.Equal(new[] { "Cross Source!A2" }, crossSecond.Dependencies);
            Assert.Equal(new[] { "Cross Source!A2" }, crossSecond.FormulaDependencies);
            Assert.Equal(4, graph.EdgeCount);
            Assert.True(graph.HasCircularReferences);
            Assert.Throws<InvalidOperationException>(() => document.InspectFormulas().EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesStructuredColumnRanges() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Structured Column Ranges");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(1, 3).SetValue("C");
            sheet.Cell(2, 1).SetFormula("SUM(Sales[[B]:[C]])");
            sheet.Cell(2, 2).SetFormula("A2");
            sheet.Cell(2, 3).SetValue(1d);
            sheet.Cell(2, 5).SetFormula("SUM(Sales[[#Data],[B]:[C]])");
            sheet.AddTable("A1:C2", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode implicitData = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Column Ranges", "A2"));
            ExcelFormulaDependencyNode explicitData = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Column Ranges", "E2"));

            Assert.Equal(new[] { "Structured Column Ranges!B2:C2" }, implicitData.Dependencies);
            Assert.Equal(new[] { "Structured Column Ranges!B2" }, implicitData.FormulaDependencies);
            Assert.Equal(new[] { "Structured Column Ranges!B2:C2" }, explicitData.Dependencies);
            Assert.Equal(new[] { "Structured Column Ranges!B2" }, explicitData.FormulaDependencies);
            Assert.True(implicitData.IsCircular);
            Assert.Throws<InvalidOperationException>(() => document.InspectFormulas().EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesEscapedStructuredReferenceBrackets() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Escaped Structured Brackets");
            sheet.Cell(1, 1).SetValue("Cost [ old");
            sheet.Cell(1, 2).SetValue("Cost ] new");
            sheet.Cell(2, 1).SetValue(1d);
            sheet.Cell(2, 2).SetValue(2d);
            sheet.Cell(1, 4).SetFormula("SUM(BracketData[Cost '[ old])");
            sheet.Cell(2, 4).SetFormula("SUM(BracketData[Cost '] new])");
            sheet.AddTable("A1:B2", hasHeader: true, name: "BracketData", style: ExcelTableStyle.TableStyleMedium2);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode escapedOpening = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Escaped Structured Brackets", "D1"));
            ExcelFormulaDependencyNode escapedClosing = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Escaped Structured Brackets", "D2"));

            Assert.Equal(new[] { "Escaped Structured Brackets!A2" }, escapedOpening.Dependencies);
            Assert.Equal(new[] { "Escaped Structured Brackets!B2" }, escapedClosing.Dependencies);
            Assert.Equal(2, document.Calculate());
            Assert.Contains(document.InspectFormulas().Formulas, formula => formula.CellReference == "D1" && formula.CachedValue == "1");
            Assert.Contains(document.InspectFormulas().Formulas, formula => formula.CellReference == "D2" && formula.CachedValue == "2");
        }

        [Fact]
        public void Test_FormulaEvaluator_ResolvesCurrentRowStructuredReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Structured Evaluation");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(1, 3).SetValue("At");
            sheet.Cell(1, 4).SetValue("Explicit");
            sheet.Cell(1, 5).SetValue("Unqualified At");
            sheet.Cell(1, 6).SetValue("Unqualified Plain");
            sheet.Cell(1, 7).SetValue("Unqualified Compare");
            sheet.Cell(1, 8).SetValue("Unqualified Direct");
            sheet.Cell(2, 1).SetValue(10d);
            sheet.Cell(2, 2).SetValue(20d);
            sheet.Cell(2, 3).SetFormula("Sales[@B]+1");
            sheet.Cell(2, 4).SetFormula("SUM(Sales[[#This Row],[B]])");
            sheet.Cell(2, 5).SetFormula("[@B]+1");
            sheet.Cell(2, 6).SetFormula("SUM([B])");
            sheet.Cell(2, 7).SetFormula("IF([@B]>10,1,0)");
            sheet.Cell(2, 8).SetFormula("[@B]");
            sheet.Cell(4, 8).SetFormula("[@B]");
            sheet.AddTable("A1:H2", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);

            int calculated = document.Calculate();
            Assert.True(sheet.TryGetCachedFormulaValue(2, 3, out string? atValue));
            Assert.Equal("21", atValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 4, out string? explicitValue));
            Assert.Equal("20", explicitValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 5, out string? unqualifiedAtValue));
            Assert.Equal("21", unqualifiedAtValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 6, out string? unqualifiedPlainValue));
            Assert.Equal("20", unqualifiedPlainValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 7, out string? unqualifiedCompareValue));
            Assert.Equal("1", unqualifiedCompareValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 8, out string? unqualifiedDirectValue));
            Assert.Equal("20", unqualifiedDirectValue);
            Assert.False(sheet.TryGetCachedFormulaValue(4, 8, out _));
            Assert.Equal(6, calculated);
        }

        [Fact]
        public void Test_FormulaStructuredReferences_ResolveWhitespaceAndCombinedItems() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Structured Combined Items");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(1, 3).SetValue("C");
            sheet.Cell(2, 1).SetValue(10d);
            sheet.Cell(2, 2).SetFormula("1+1");
            sheet.Cell(2, 3).SetValue(30d);
            sheet.Cell(3, 1).SetValue(1d);
            sheet.Cell(3, 2).SetValue(5d);
            sheet.Cell(3, 3).SetValue(3d);
            sheet.Cell(1, 5).SetFormula("SUM(Sales[[#Data], [A] : [C]])");
            sheet.Cell(2, 5).SetFormula("SUM(Sales[[#Headers],[#Data],[B]])");
            sheet.Cell(3, 5).SetFormula("SUM(Sales[[#Data],[#Totals],[B]])");
            sheet.AddTable("A1:C3", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(
                document.WorkbookPartRoot.WorksheetParts.SelectMany(part => part.TableDefinitionParts)).Table;
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            table.Save();

            Assert.Equal(4, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 5, out string? whitespaceValue));
            Assert.Equal("42", whitespaceValue);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 5, out string? headerDataValue));
            Assert.Equal("2", headerDataValue);
            Assert.True(sheet.TryGetCachedFormulaValue(3, 5, out string? dataTotalsValue));
            Assert.Equal("7", dataTotalsValue);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode whitespace = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Combined Items", "E1"));
            ExcelFormulaDependencyNode headerData = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Combined Items", "E2"));
            ExcelFormulaDependencyNode dataTotals = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Combined Items", "E3"));

            Assert.Equal(new[] { "Structured Combined Items!A2:C2" }, whitespace.Dependencies);
            Assert.Equal(new[] { "Structured Combined Items!B2" }, whitespace.FormulaDependencies);
            Assert.Equal(new[] { "Structured Combined Items!B1:B2" }, headerData.Dependencies);
            Assert.Equal(new[] { "Structured Combined Items!B2" }, headerData.FormulaDependencies);
            Assert.Equal(new[] { "Structured Combined Items!B2:B3" }, dataTotals.Dependencies);
            Assert.Equal(new[] { "Structured Combined Items!B2" }, dataTotals.FormulaDependencies);
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesUnqualifiedCurrentRowStructuredReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Unqualified Current Row");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(1, 3).SetValue("Sales Amount");
            sheet.Cell(1, 4).SetValue("% Commission");
            sheet.Cell(1, 5).SetValue("Book.xlsx");
            sheet.Cell(2, 1).SetFormula("[@B]");
            sheet.Cell(2, 2).SetFormula("[[#This Row],[A]]");
            sheet.Cell(2, 3).SetFormula("[% Commission]");
            sheet.Cell(2, 4).SetFormula("[@[Sales Amount]]");
            sheet.Cell(2, 6).SetFormula("[@B]");
            sheet.Cell(3, 1).SetFormula("UnknownTable[[#This Row],[B]]");
            sheet.Cell(3, 2).SetFormula("[@B]Sheet1!A1");
            sheet.Cell(3, 3).SetValue(1d);
            sheet.Cell(3, 4).SetValue(2d);
            sheet.Cell(3, 5).SetFormula("[Book.xlsx]Sheet1!RemoteName");
            sheet.AddTable(
                "A1:E3",
                hasHeader: true,
                name: "UnqualifiedCurrentRowData",
                style: ExcelTableStyle.TableStyleMedium2);

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode first = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "A2"));
            Assert.Equal(new[] { "Unqualified Current Row!B2" }, first.Dependencies);
            Assert.Equal(new[] { "Unqualified Current Row!B2" }, first.FormulaDependencies);
            ExcelFormulaDependencyNode second = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "B2"));
            Assert.Equal(new[] { "Unqualified Current Row!A2" }, second.Dependencies);
            Assert.Equal(new[] { "Unqualified Current Row!A2" }, second.FormulaDependencies);
            ExcelFormulaDependencyNode spaced = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "C2"));
            Assert.Equal(new[] { "Unqualified Current Row!D2" }, spaced.Dependencies);
            Assert.Equal(new[] { "Unqualified Current Row!D2" }, spaced.FormulaDependencies);
            ExcelFormulaDependencyNode nested = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "D2"));
            Assert.Equal(new[] { "Unqualified Current Row!C2" }, nested.Dependencies);
            Assert.Equal(new[] { "Unqualified Current Row!C2" }, nested.FormulaDependencies);
            ExcelFormulaDependencyNode outsideTable = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "F2"));
            Assert.Empty(outsideTable.Dependencies);
            Assert.Empty(outsideTable.FormulaDependencies);
            ExcelFormulaDependencyNode unresolvedTable = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "A3"));
            Assert.Empty(unresolvedTable.Dependencies);
            Assert.Empty(unresolvedTable.FormulaDependencies);
            ExcelFormulaDependencyNode externalReference = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "B3"));
            Assert.Empty(externalReference.Dependencies);
            Assert.Empty(externalReference.FormulaDependencies);
            ExcelFormulaDependencyNode externalWorkbook = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Unqualified Current Row", "E3"));
            Assert.Empty(externalWorkbook.Dependencies);
            Assert.Empty(externalWorkbook.FormulaDependencies);
            Assert.Equal(4, graph.EdgeCount);
            Assert.Equal(2, graph.CircularReferenceCount);
            Assert.True(graph.HasCircularReferences);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_IgnoresLexicallyScopedNames() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Lexical");
            sheet.Cell(1, 1).SetFormula("LET(Input,1,Input+1)");
            sheet.Cell(1, 2).SetFormula("A1");
            sheet.Cell(2, 1).SetFormula("Input+LET(Input,1,Input+1)");
            sheet.Cell(2, 2).SetFormula("LAMBDA(Input,Input+1)(1)");
            sheet.Cell(3, 1).SetFormula("LET(Input,{1,2},SUM(Input))");
            sheet.Cell(1, 4).SetValue("Amount");
            sheet.Cell(1, 5).SetValue("Other");
            sheet.Cell(2, 4).SetValue(1d);
            sheet.Cell(2, 5).SetValue(2d);
            sheet.AddTable("D1:E2", hasHeader: true, name: "LexicalData", style: ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("LET(Input,LexicalData[[#Headers],[Amount]],Input)");
            ExcelSheet quotedQualifier = document.AddSheet("O'A,B)");
            quotedQualifier.Cell(1, 1).SetValue(5d);
            sheet.Cell(5, 1).SetFormula("LET(Input,'O''A,B)'!A1,Input)");
            sheet.Cell(6, 1).SetFormula("LET(LexicalData,1,SUM(LexicalData[Amount]))");
            document.SetNamedRange("Input", "Lexical!B1", save: false);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode let = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A1"));
            Assert.Empty(let.Dependencies);
            Assert.Empty(let.FormulaDependencies);
            ExcelFormulaDependencyNode mixed = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A2"));
            Assert.Equal(new[] { "Lexical!B1" }, mixed.Dependencies);
            Assert.Equal(new[] { "Lexical!B1" }, mixed.FormulaDependencies);
            ExcelFormulaDependencyNode lambda = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "B2"));
            Assert.Empty(lambda.Dependencies);
            ExcelFormulaDependencyNode array = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A3"));
            Assert.Empty(array.Dependencies);
            ExcelFormulaDependencyNode structured = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A4"));
            Assert.Equal(new[] { "Lexical!D1" }, structured.Dependencies);
            ExcelFormulaDependencyNode quoted = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A5"));
            Assert.Equal(new[] { "O'A,B)!A1" }, quoted.Dependencies);
            ExcelFormulaDependencyNode tableCollision = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Lexical", "A6"));
            Assert.Equal(new[] { "Lexical!D2" }, tableCollision.Dependencies);
            Assert.False(graph.HasCircularReferences);
        }

        [Fact]
        public void Test_FormulaDependencyInspection_EvaluatesCrossSheetDependenciesOnOwningSheet() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet source = document.AddSheet("Source");
            source.Cell(1, 1).SetFormula("Other!C2");
            ExcelSheet other = document.AddSheet("Other");
            other.Cell(1, 1).SetValue("Amount");
            other.Cell(1, 2).SetValue("Label");
            other.Cell(1, 3).SetValue("Result");
            other.Cell(2, 1).SetValue(2d);
            other.Cell(2, 2).SetValue("Row");
            other.Cell(2, 3).SetFormula("OtherSales[@Amount]+1");
            other.AddTable("A1:C2", hasHeader: true, name: "OtherSales", style: ExcelTableStyle.TableStyleMedium2);
            Assert.Equal(1, other.RecalculateSupportedFormulas());

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaCellInfo sourceFormula = Assert.Single(inspection.Formulas, formula =>
                formula.SheetName == "Source" && formula.CellReference == "A1");
            Assert.Empty(sourceFormula.DependencyIssues);
            inspection.EnsureNoDependencyIssues();
        }

        [Fact]
        public void Test_FormulaEvaluator_StopsAtConfiguredDependencyDepth() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Budget.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Depth");
                sheet.Cell(4, 1).SetValue(1d);
                sheet.Cell(3, 1).SetFormula("A4+1");
                sheet.Cell(2, 1).SetFormula("A3+1");
                sheet.Cell(1, 1).SetFormula("A2+1");

                ExcelSheet dependenciesFirst = document.AddSheet("DependenciesFirst");
                dependenciesFirst.Cell(1, 1).SetValue(1d);
                dependenciesFirst.Cell(2, 1).SetFormula("A1+1");
                dependenciesFirst.Cell(3, 1).SetFormula("A2+1");
                dependenciesFirst.Cell(4, 1).SetFormula("A3+1");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
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
                ExcelSheet named = document.AddSheet("Named Circular");
                named.Cell(1, 1).SetFormula("Loop+1");
                document.SetNamedRange("Loop", "'Named Circular'!A1", save: false);
                document.SetNamedRange("Amount", "'Named Circular'!A1", save: false);
                document.SetNamedRange("SUM", "'Named Circular'!A1", save: false);

                ExcelSheet structured = document.AddSheet("Structured");
                structured.Cell(1, 1).SetValue("Amount");
                structured.Cell(2, 1).SetFormula("SUM(Sales[Amount])");
                structured.Cell(1, 2).SetValue("A1");
                structured.Cell(2, 2).SetValue(2d);
                structured.Cell(2, 3).SetFormula("SUM(Sales[A1])");
                structured.AddTable("A1:B2", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);

                ExcelSheet sales = document.AddSheet("Sales");
                sales.Cell(1, 1).SetValue(5d);
                ExcelSheet salesData = document.AddSheet("Sales Data");
                salesData.Cell(1, 3).SetValue(7d);
                salesData.SetNamedRange("SharedInput", "C1", save: false);
                ExcelSheet apostropheSales = document.AddSheet("O'Sales");
                apostropheSales.Cell(1, 3).SetValue(8d);
                apostropheSales.SetNamedRange("EscapedInput", "C1", save: false);

                ExcelSheet tokens = document.AddSheet("Tokens");
                tokens.Cell(1, 1).SetValue(100d);
                tokens.Cell(1, 2).SetFormula("SUM(A1)");
                tokens.Cell(1, 3).SetFormula("LOG10(A1)");
                tokens.Cell(1, 4).SetFormula("Sales!A1+1");
                tokens.Cell(1, 5).SetFormula("'Sales Data'!C1");
                tokens.Cell(1, 6).SetFormula("'Sales Data'!SharedInput+1");
                tokens.Cell(1, 7).SetFormula("'O''Sales'!EscapedInput+1");
                document.SetNamedRange("Sales", "Tokens!A1", save: false);

                ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
                Assert.Equal(9, graph.NodeCount);
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

                ExcelFormulaDependencyNode cellLikeColumnNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Structured", "C2"));
                Assert.Equal(new[] { "Structured!B2" }, cellLikeColumnNode.Dependencies);
                Assert.Empty(cellLikeColumnNode.FormulaDependencies);

                ExcelFormulaDependencyNode sumNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "B1"));
                Assert.Equal(new[] { "Tokens!A1" }, sumNode.Dependencies);
                Assert.Empty(sumNode.FormulaDependencies);

                ExcelFormulaDependencyNode logNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "C1"));
                Assert.Equal(new[] { "Tokens!A1" }, logNode.Dependencies);
                Assert.Empty(logNode.FormulaDependencies);

                ExcelFormulaDependencyNode qualifiedNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "D1"));
                Assert.Equal(new[] { "Sales!A1" }, qualifiedNode.Dependencies);
                Assert.Empty(qualifiedNode.FormulaDependencies);

                ExcelFormulaDependencyNode quotedQualifierNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "E1"));
                Assert.Equal(new[] { "Sales Data!C1" }, quotedQualifierNode.Dependencies);
                Assert.Empty(quotedQualifierNode.FormulaDependencies);

                ExcelFormulaDependencyNode scopedAliasNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "F1"));
                Assert.Equal(new[] { "Sales Data!C1" }, scopedAliasNode.Dependencies);
                Assert.Empty(scopedAliasNode.FormulaDependencies);

                ExcelFormulaDependencyNode escapedQualifierNode = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "G1"));
                Assert.Equal(new[] { "O'Sales!C1" }, escapedQualifierNode.Dependencies);
                Assert.Empty(escapedQualifierNode.FormulaDependencies);

                Assert.Throws<InvalidOperationException>(() => document.InspectFormulas().EnsureNoDependencyIssues());
            }
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesDefinedNamesWithQuotedCommas() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Sales, East");
            sheet.Cell(1, 1).SetFormula("B1");
            sheet.Cell(1, 2).SetFormula("CommaAlias");
            document.SetNamedRange("CommaAlias", "'Sales, East'!A1", save: false);

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode aliasNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Sales, East", "B1"));

            Assert.Equal(new[] { "Sales, East!A1" }, aliasNode.Dependencies);
            Assert.Equal(new[] { "Sales, East!A1" }, aliasNode.FormulaDependencies);
            Assert.Equal(2, graph.EdgeCount);
            Assert.Equal(
                new[] { "Sales, East!A1", "Sales, East!B1" },
                Assert.Single(graph.CircularReferences).References);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesWholeRowAndColumnDefinedNames() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet columns = document.AddSheet("Whole Column Name");
            columns.Cell(1, 1).SetFormula("SUM(AllB)");
            columns.Cell(1, 2).SetFormula("A1");
            ExcelSheet rows = document.AddSheet("Whole Row Name");
            rows.Cell(1, 1).SetFormula("SUM(AllRows)");
            rows.Cell(2, 1).SetFormula("A1");
            document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            document.WorkbookRoot.DefinedNames.Append(
                new DefinedName("'Whole Column Name'!B:B") { Name = "AllB" },
                new DefinedName("'Whole Row Name'!2:2") { Name = "AllRows" });

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode columnNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Whole Column Name", "A1"));
            ExcelFormulaDependencyNode rowNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Whole Row Name", "A1"));

            Assert.Equal(new[] { "Whole Column Name!B1:B1048576" }, columnNode.Dependencies);
            Assert.Equal(new[] { "Whole Column Name!B1" }, columnNode.FormulaDependencies);
            Assert.Equal(new[] { "Whole Row Name!A2:XFD2" }, rowNode.Dependencies);
            Assert.Equal(new[] { "Whole Row Name!A2" }, rowNode.FormulaDependencies);
            Assert.Equal(4, graph.EdgeCount);
            Assert.Equal(2, graph.CircularReferenceCount);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesBackslashTableAndStructuredDefinedName() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Backslash Table");
            sheet.Cell(1, 1).SetValue("Amount");
            sheet.Cell(2, 1).SetFormula("C2");
            sheet.Cell(2, 3).SetFormula("SUM(TableAlias)");
            sheet.Cell(1, 4).SetFormula("SUM(\\Sales[Amount])");
            sheet.AddTable("A1:A2", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(
                document.WorkbookPartRoot.WorksheetParts.SelectMany(part => part.TableDefinitionParts)).Table;
            table.Name = "\\Sales";
            table.DisplayName = "\\Sales";
            document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            document.WorkbookRoot.DefinedNames.Append(
                new DefinedName("CurrentAmount") { Name = "TableAlias" },
                new DefinedName("\\Sales[[#This Row],[Amount]]") { Name = "CurrentAmount" });

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode aliasNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Backslash Table", "C2"));
            ExcelFormulaDependencyNode directNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Backslash Table", "D1"));

            Assert.Equal(new[] { "Backslash Table!A2" }, aliasNode.Dependencies);
            Assert.Equal(new[] { "Backslash Table!A2" }, aliasNode.FormulaDependencies);
            Assert.Equal(new[] { "Backslash Table!A2" }, directNode.Dependencies);
            Assert.Equal(new[] { "Backslash Table!A2" }, directNode.FormulaDependencies);
            Assert.Equal(3, graph.EdgeCount);
            Assert.Equal(
                new[] { "Backslash Table!A2", "Backslash Table!C2" },
                Assert.Single(graph.CircularReferences).References);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesChainedDefinedNamesAndGuardsNameCycles() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Chained Names");
            sheet.Cell(1, 1).SetFormula("B1");
            sheet.Cell(1, 2).SetFormula("Alias");
            sheet.Cell(1, 3).SetFormula("First");
            document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            document.WorkbookRoot.DefinedNames.Append(
                new DefinedName("'Chained Names'!A1") { Name = "Base" },
                new DefinedName("Base") { Name = "Alias" },
                new DefinedName("Second") { Name = "First" },
                new DefinedName("First") { Name = "Second" });

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode aliasNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Chained Names", "B1"));
            ExcelFormulaDependencyNode cyclicNameNode = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Chained Names", "C1"));

            Assert.Equal(new[] { "Chained Names!A1" }, aliasNode.Dependencies);
            Assert.Equal(new[] { "Chained Names!A1" }, aliasNode.FormulaDependencies);
            Assert.Empty(cyclicNameNode.Dependencies);
            Assert.Empty(cyclicNameNode.FormulaDependencies);
            Assert.Equal(2, graph.EdgeCount);
            Assert.True(graph.HasCircularReferences);
            Assert.Equal(new[] { "Chained Names!A1", "Chained Names!B1" }, Assert.Single(graph.CircularReferences).References);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ExpandsDefinedNameRangeEndpoints() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Named Range Endpoints");
            sheet.Cell(1, 1).SetValue(1d);
            sheet.Cell(1, 2).SetFormula("D1");
            sheet.Cell(1, 3).SetValue(3d);
            sheet.Cell(1, 4).SetFormula("SUM(Start:End)");
            sheet.Cell(1, 5).SetFormula("SUM((Start):(End))");
            sheet.Cell(1, 6).SetFormula("SUM(Start:C1)");
            sheet.Cell(1, 7).SetFormula("SUM(A1:End)");
            document.SetNamedRange("Start", "'Named Range Endpoints'!A1", save: false);
            document.SetNamedRange("End", "'Named Range Endpoints'!C1", save: false);

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaDependencyGraph graph = inspection.DependencyGraph;
            ExcelFormulaDependencyNode range = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Named Range Endpoints", "D1"));

            Assert.Equal(new[] { "Named Range Endpoints!A1:C1" }, range.Dependencies);
            Assert.Equal(new[] { "Named Range Endpoints!B1" }, range.FormulaDependencies);
            ExcelFormulaDependencyNode parenthesizedRange = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Named Range Endpoints", "E1"));
            Assert.Equal(new[] { "Named Range Endpoints!A1:C1" }, parenthesizedRange.Dependencies);
            Assert.Equal(new[] { "Named Range Endpoints!B1" }, parenthesizedRange.FormulaDependencies);
            ExcelFormulaDependencyNode namedStartRange = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Named Range Endpoints", "F1"));
            Assert.Equal(new[] { "Named Range Endpoints!A1:C1" }, namedStartRange.Dependencies);
            Assert.Equal(new[] { "Named Range Endpoints!B1" }, namedStartRange.FormulaDependencies);
            ExcelFormulaDependencyNode namedEndRange = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Named Range Endpoints", "G1"));
            Assert.Equal(new[] { "Named Range Endpoints!A1:C1" }, namedEndRange.Dependencies);
            Assert.Equal(new[] { "Named Range Endpoints!B1" }, namedEndRange.FormulaDependencies);
            Assert.Equal(5, graph.EdgeCount);
            Assert.True(graph.HasCircularReferences);
            Assert.Equal(
                new[] { "Named Range Endpoints!B1", "Named Range Endpoints!D1" },
                Assert.Single(graph.CircularReferences).References);
            Assert.Throws<InvalidOperationException>(() => inspection.EnsureNoDependencyIssues());
        }

        [Fact]
        public void Test_FormulaDependencyGraph_BoundsAcyclicDefinedNameExpansion() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Bounded Names");
            sheet.Cell(1, 1).SetValue(1d);
            sheet.Cell(1, 2).SetFormula("Name000");
            document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            for (int index = 0; index < 256; index++) {
                document.WorkbookRoot.DefinedNames.Append(new DefinedName($"Name{index + 1:000}") {
                    Name = $"Name{index:000}"
                });
            }
            document.WorkbookRoot.DefinedNames.Append(new DefinedName("'Bounded Names'!A1") { Name = "Name256" });

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode node = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Bounded Names", "B1"));

            Assert.Empty(node.Dependencies);
            Assert.Empty(node.FormulaDependencies);
            Assert.False(graph.HasCircularReferences);
        }

        [Fact]
        public void Test_FormulaEvaluator_ClearsPreExistingCacheWhenLowerDepthBlocksFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.PreExistingCache.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Depth");
                sheet.Cell(1, 1).SetValue(1d);
                sheet.Cell(2, 1).SetFormula("A1+1");
                sheet.Cell(3, 1).SetFormula("A2+1");
                sheet.Cell(4, 1).SetFormula("A3+1");

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
        public void Test_FormulaEvaluator_DoesNotConsumeDepthBlockedValuesInInfoFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.InfoFunction.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Depth");
            sheet.Cell(1, 1).SetFormula("ISBLANK(A2)");
            sheet.Cell(2, 1).SetFormula("1+1");
            document.Calculation.MaximumDependencyDepth = 1;

            Assert.Equal(1, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? cached));
            Assert.Equal("2", cached);
        }

        [Fact]
        public void Test_FormulaEvaluator_ChargesDependencyDepthForARangeReadEarlierInThePass() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Depth");
            sheet.Cell(1, 1).SetValue(1d);
            sheet.Cell(2, 1).SetFormula("A1+1");
            sheet.Cell(1, 3).SetFormula("SUMIFS(A1:A2,A1:A2,\">0\")");
            sheet.Cell(2, 3).SetFormula("C3+1");
            sheet.Cell(3, 3).SetFormula("SUMIFS(A1:A2,A1:A2,\">0\")");
            document.Calculation.MaximumDependencyDepth = 2;

            document.Calculate();

            // C1 reads the range first. C2 reaches it one level deeper, where A2 exceeds the depth budget.
            Assert.True(sheet.TryGetCachedFormulaValue(1, 3, out string? first));
            Assert.Equal("3", first);
            Assert.False(sheet.TryGetCachedFormulaValue(2, 3, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(3, 3, out string? direct));
            Assert.Equal("3", direct);
        }

        [Fact]
        public void Test_FormulaEvaluator_LookupStopsReadingAtItsMatch() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Lookup");
            sheet.Cell(1, 1).SetValue("a");
            sheet.Cell(2, 1).SetFormula("B1");
            sheet.Cell(1, 2).SetFormula("MATCH(\"a\",A1:A2,0)");

            document.Calculate();

            // Reading A2 would reach B1 itself; the match in A1 ends the search before it.
            Assert.True(sheet.TryGetCachedFormulaValue(1, 2, out string? position));
            Assert.Equal("1", position);
        }

        [Fact]
        public void Test_FormulaEvaluator_EnforcesDependencyDepthAcrossWorksheets() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet source = document.AddSheet("Source");
            ExcelSheet other = document.AddSheet("Other");
            source.Cell(1, 1).SetFormula("Other!A1+1");
            other.Cell(1, 1).SetFormula("1+1");
            document.Calculation.MaximumDependencyDepth = 1;

            Assert.Equal(1, document.Calculate());
            Assert.False(source.TryGetCachedFormulaValue(1, 1, out _));
            Assert.True(other.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal("2", cached);
        }

        [Fact]
        public void Test_FormulaEvaluator_DoesNotCacheAggregatesAfterDependencyDepthGuard() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Depth");
            sheet.Cell(1, 1).SetFormula("SUBTOTAL(2,A2)");
            sheet.Cell(1, 2).SetFormula("COUNTBLANK(B2)");
            sheet.Cell(1, 3).SetFormula("COUNTIF(C2,\">0\")");
            sheet.Cell(2, 1).SetFormula("1+1");
            sheet.Cell(2, 2).SetFormula("1+1");
            sheet.Cell(2, 3).SetFormula("1+1");
            document.Calculation.MaximumDependencyDepth = 1;

            Assert.Equal(3, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
            Assert.False(sheet.TryGetCachedFormulaValue(1, 2, out _));
            Assert.False(sheet.TryGetCachedFormulaValue(1, 3, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 2, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 3, out _));
        }

        [Fact]
        public void Test_FormulaEvaluator_CountsCachedUnsupportedFormulaDependenciesTowardDepth() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Depth");
            sheet.Cell(1, 1).SetFormula("UNSUPPORTED()");
            sheet.Cell(2, 1).SetFormula("SUM(A1)");
            sheet.Cell(3, 1).SetFormula("SUM(A2)");

            Cell unsupported = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell =>
                cell.CellReference?.Value == "A1");
            unsupported.CellValue = new CellValue("1");
            unsupported.DataType = CellValues.Number;
            sheet.WorksheetPart.Worksheet.Save();
            document.Calculation.MaximumDependencyDepth = 2;

            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? unsupportedCache));
            Assert.Equal("1", unsupportedCache);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? cachedA2));
            Assert.Equal("1", cachedA2);
            Assert.False(sheet.TryGetCachedFormulaValue(3, 1, out _));
        }

        [Fact]
        public void Test_FormulaCalculationOptions_ValidateDependencyDepth() {
            var options = new ExcelCalculationOptions();

            Assert.Equal(256, options.MaximumDependencyDepth);
            Assert.Throws<ArgumentOutOfRangeException>(() => options.MaximumDependencyDepth = 0);
            Assert.Equal(256, options.MaximumDependencyDepth);
        }

        [Fact]
        public void Test_FormulaInspection_ExpandsSharedFormulaFollowers() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.Shared.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.Cell(1, 1).SetValue(1d);
                sheet.Cell(2, 1).SetValue(2d);
                sheet.Cell(1, 2).SetFormula("SUM(A1,$A1,A$1,$A$1)");
                sheet.Cell(1, 3).SetFormula("SUM(B1,$A1,B$1,$A$1)");
                sheet.Cell(2, 2).SetFormula("SUM(A2,$A2,A$1,$A$1)");
                sheet.Cell(2, 3).SetFormula("SUM(B2,$A2,B$1,$A$1)");
                Assert.Equal(4, document.Calculate());
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell[] formulaCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellFormula != null)
                    .ToArray();
                Cell master = Assert.Single(formulaCells, cell => cell.CellReference?.Value == "B1");
                master.CellFormula = new CellFormula("SUM(A1,$A1,A$1,$A$1)") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 0U,
                    Reference = "B1:C2"
                };
                foreach (Cell follower in formulaCells.Where(cell => cell != master)) {
                    follower.CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0U
                    };
                }
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets[0];
                Assert.Equal("SUM(A1,$A1,A$1,$A$1)", sheet.GetFormulaText(1, 2));
                Assert.Equal("SUM(B1,$A1,B$1,$A$1)", sheet.GetFormulaText(1, 3));
                Assert.Equal("SUM(A2,$A2,A$1,$A$1)", sheet.GetFormulaText(2, 2));
                Assert.Equal("SUM(B2,$A2,B$1,$A$1)", sheet.GetFormulaText(2, 3));

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(4, inspection.Formulas.Count);
                Assert.All(inspection.Formulas, formula => Assert.True(formula.IsSupportedByOfficeIMO));
                Assert.Equal(new[] { "Shared!A1" }, Assert.Single(inspection.Formulas, formula => formula.CellReference == "B1").Dependencies);
                Assert.Contains("Shared!B2", Assert.Single(inspection.Formulas, formula => formula.CellReference == "C2").Dependencies);

                ExcelWorkbookInfo snapshot = document.Inspect();
                ExcelCellInfo snapshotFollower = Assert.Single(
                    Assert.Single(snapshot.Worksheets).Cells,
                    cell => cell.Row == 2 && cell.Column == 3);
                Assert.Equal("SUM(B2,$A2,B$1,$A$1)", snapshotFollower.Formula);

                foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xls, ExcelFileFormat.Xlsb }) {
                    byte[] binary = document.ToBytes(format);
                    using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(binary, writable: false));
                    Assert.Equal(format, converted.SourceFormat);
                    Assert.Equal("SUM(B2,$A2,B$1,$A$1)", converted["Shared"].GetFormulaText(2, 3));
                }

                Assert.Equal(4, document.Calculate());
                Assert.True(sheet.TryGetCachedFormulaValue(1, 3, out string? cachedC1));
                Assert.Equal("10", cachedC1);
                Assert.True(sheet.TryGetCachedFormulaValue(2, 3, out string? cachedC2));
                Assert.Equal("13", cachedC2);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FormulaText_ResolvesSharedFormulaFollowers() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.SharedFormulaText.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared Formula Text");
                sheet.Cell(1, 1).SetValue(1d);
                sheet.Cell(2, 1).SetValue(2d);
                sheet.Cell(1, 2).SetFormula("A1+1");
                sheet.Cell(2, 2).SetFormula("A2+1");
                sheet.Cell(1, 3).SetFormula("FORMULATEXT(B2)");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = Assert.Single(
                    worksheet.Descendants<Cell>(),
                    cell => cell.CellReference?.Value == "B1");
                Cell follower = Assert.Single(
                    worksheet.Descendants<Cell>(),
                    cell => cell.CellReference?.Value == "B2");
                master.CellFormula = new CellFormula("A1+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 11U,
                    Reference = "B1:B2"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 11U
                };
                worksheet.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded["Shared Formula Text"];
            Assert.Equal("A2+1", loadedSheet.GetFormulaText(2, 2));
            Assert.Equal(3, loaded.Calculate());
            Assert.True(loadedSheet.TryGetCachedFormulaValue(1, 3, out string? formulaText));
            Assert.Equal("=A2+1", formulaText);
            Assert.Empty(loaded.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaInspection_ExpandsSharedWholeRowColumnSpillAndThreeDimensionalReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.SharedReferenceKinds.xlsx");
            const string masterFormula = "SUM(A:A,$A:$A,1:1,$1:$1,A1#,$A1#,A$1#,$A$1#,A1%)+LOG10(A1)+LOG10 (A1)+FOO1 (A1)+SUM(A1 (A1))+SUM(A1 (A1:A2))+SUM(A1 (A1,B2))+SUM(A1 ((A1,B2),C3))+Q1:Q4!A1";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Q1");
                document.AddSheet("Q4");
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.Cell(1, 5).SetFormula(masterFormula);
                sheet.Cell(1, 6).SetFormula(masterFormula);
                sheet.Cell(2, 5).SetFormula(masterFormula);
                sheet.Cell(2, 6).SetFormula(masterFormula);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.Worksheet.Descendants<Cell>().Any(cell => cell.CellReference?.Value == "E1"))
                    .Worksheet;
                Cell[] formulaCells = worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellFormula != null)
                    .ToArray();
                Cell master = Assert.Single(formulaCells, cell => cell.CellReference?.Value == "E1");
                master.CellFormula = new CellFormula(masterFormula) {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 7U,
                    Reference = "E1:F2"
                };
                foreach (Cell follower in formulaCells.Where(cell => cell != master)) {
                    follower.CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 7U
                    };
                }
                worksheet.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet shared = loaded["Shared"];
            Assert.Equal(masterFormula, shared.GetFormulaText(1, 5));
            // Built-in functions remain stable while cell-like tokens followed by
            // parenthesized references participate in whitespace intersections.
            Assert.Equal(
                "SUM(B:B,$A:$A,2:2,$1:$1,B2#,$A2#,B$1#,$A$1#,B2%)+LOG10(B2)+LOG10 (B2)+FOP2 (B2)+SUM(B2 (B2))+SUM(B2 (B2:B3))+SUM(B2 (B2,C3))+SUM(B2 ((B2,C3),D4))+Q1:Q4!B2",
                shared.GetFormulaText(2, 6));
        }

#if EXCEL_PERFORMANCE_EVIDENCE
        [Trait("Category", "Performance")]
        [Trait("Category", "ExcelPerformanceEvidence")]
#endif
        [Fact]
        public void Test_FormulaInspection_DoesNotExpandOversizedSharedFormulaFollowers() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Shared");
            Worksheet worksheet = sheet.WorksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
#if EXCEL_PERFORMANCE_EVIDENCE
            const int followerCount = 1_000;
#else
            const int followerCount = 2;
#endif
            string masterFormula = new string('A', 8_193);
            sheetData.RemoveAllChildren<Row>();
            for (int rowIndex = 1; rowIndex <= followerCount; rowIndex++) {
                var cell = new Cell { CellReference = "A" + rowIndex };
                cell.CellFormula = rowIndex == 1
                    ? new CellFormula(masterFormula) {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 19U,
                        Reference = "A1:A" + followerCount
                    }
                    : new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 19U
                    };
                sheetData.Append(new Row(cell) { RowIndex = (uint)rowIndex });
            }
            worksheet.Save();

#if EXCEL_PERFORMANCE_EVIDENCE
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
#endif
            Assert.Equal(string.Empty, sheet.GetFormulaText(followerCount, 1));
            ExcelFormulaInspection inspection = document.InspectFormulas();
#if EXCEL_PERFORMANCE_EVIDENCE
            stopwatch.Stop();
#endif

            Assert.Equal(followerCount, inspection.Formulas.Count);
#if EXCEL_PERFORMANCE_EVIDENCE
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), stopwatch.Elapsed.ToString());
#endif
        }

        [Fact]
        public void Test_FormulaInspection_IgnoresDefinedNamesScopedToMalformedSheets() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet valid = document.AddSheet("Valid");
            valid.Cell(1, 1).SetFormula("Broken!ScopedValue");

            uint malformedSheetIndex = (uint)document.WorkbookRoot.Sheets!.Elements<Sheet>().Count();
            document.WorkbookRoot.Sheets.Append(new Sheet {
                Name = "Broken",
                SheetId = malformedSheetIndex + 1U
            });
            document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            document.WorkbookRoot.DefinedNames.Append(new DefinedName("Valid!$B$1") {
                Name = "ScopedValue",
                LocalSheetId = malformedSheetIndex
            });
            document.WorkbookRoot.Save();

            ExcelFormulaInspection inspection = document.InspectFormulas();
            ExcelFormulaCellInfo formula = Assert.Single(inspection.Formulas);
            Assert.Equal("A1", formula.CellReference);
            Assert.Empty(formula.Dependencies);
        }

        [Fact]
        public void Test_FormulaDependencyGraph_IgnoresDefinedNamesInsideErrorLiterals() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Errors");
            sheet.Cell(1, 1).SetFormula("#NAME?");
            sheet.Cell(1, 2).SetFormula("A1+1");
            document.SetNamedRange("NAME", "Errors!B1", save: false);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode error = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Errors", "A1"));
            Assert.Empty(error.Dependencies);
            Assert.Empty(error.FormulaDependencies);
            ExcelFormulaDependencyNode dependent = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Errors", "B1"));
            Assert.Equal(new[] { "Errors!A1" }, dependent.Dependencies);
            Assert.Equal(new[] { "Errors!A1" }, dependent.FormulaDependencies);
            Assert.Equal(1, graph.EdgeCount);
            Assert.False(graph.HasCircularReferences);
        }

        [Fact]
        public void Test_FormulaDependencyGraph_DoesNotTreatExternalOrThreeDimensionalReferencesAsLocal() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet data = document.AddSheet("Data");
            data.Cell(1, 1).SetValue(10d);
            ExcelSheet first = document.AddSheet("First");
            first.Cell(1, 1).SetValue(20d);
            ExcelSheet last = document.AddSheet("Last");
            last.Cell(1, 1).SetValue(30d);
            document.AddSheet("Q1");
            document.AddSheet("Q4");
            ExcelSheet tokens = document.AddSheet("Tokens");
            tokens.Cell(1, 1).SetFormula("[Other.xlsx]Data!A1+First:Last!A1+'First:Last'!A1+Q1:Q4!A1");
            document.SetNamedRange("Data", "Data!A1", save: false);
            document.SetNamedRange("First", "First!A1", save: false);

            ExcelFormulaDependencyGraph graph = document.InspectFormulas().DependencyGraph;
            ExcelFormulaDependencyNode node = Assert.IsType<ExcelFormulaDependencyNode>(graph.FindNode("Tokens", "A1"));
            Assert.Empty(node.Dependencies);
            Assert.Empty(node.FormulaDependencies);
            Assert.Equal(0, graph.EdgeCount);
            Assert.False(graph.HasCircularReferences);
        }
    }
}
