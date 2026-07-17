using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using TableStyle = OfficeIMO.Excel.TableStyle;

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

                ExcelFeatureFinding dependencyFinding = document.InspectFeatures().Features
                    .Single(feature => feature.Name == "Formula dependency issues");
                Assert.Contains(dependencyFinding.Details, detail =>
                    detail.Contains("Circular reference: Circular!A1 -> Circular!B1", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesWholeRowAndColumnReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet columns = document.AddWorksheet("Whole Columns");
            columns.CellFormula(1, 1, "SUM(B:B)");
            columns.CellFormula(1, 2, "A1");
            ExcelSheet rows = document.AddWorksheet("Whole Rows");
            rows.CellFormula(1, 1, "SUM(2:2)");
            rows.CellFormula(2, 1, "A1");

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
        public void Test_FormulaDependencyGraph_RespectsRangeIntersections() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Intersections");
            sheet.CellFormula(1, 1, "D1");
            sheet.CellFormula(1, 2, "D3");
            sheet.CellFormula(2, 2, "1+1");
            sheet.CellFormula(1, 4, "SUM(A1:B2 B2:C3)");
            sheet.CellFormula(2, 4, "SUM(Data B2:C3)");
            sheet.CellFormula(3, 4, "SUM(A1 (A1:B2))");
            sheet.CellFormula(4, 4, "SUM((A1:B2) B2:C3)");
            sheet.CellFormula(5, 4, "SUM((A1:B2) (B2:C3))");
            sheet.CellFormula(6, 4, "SUM((A1:B2 B2:C3) B2:D2)");
            document.SetNamedRange("Data", "Intersections!A1:B2", save: false);

            ExcelSheet structuredSheet = document.AddWorksheet("Structured Intersections");
            structuredSheet.CellValue(1, 1, "Amount");
            structuredSheet.CellValue(1, 2, "Other");
            structuredSheet.CellFormula(2, 1, "1+1");
            structuredSheet.CellValue(2, 2, 3d);
            structuredSheet.CellFormula(1, 4, "SUM(IntersectionData[Amount] A2:B2)");
            structuredSheet.AddTable(
                "A1:B2",
                hasHeader: true,
                name: "IntersectionData",
                style: TableStyle.TableStyleMedium2);

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
            ExcelFormulaDependencyNode structuredIntersection = Assert.IsType<ExcelFormulaDependencyNode>(
                graph.FindNode("Structured Intersections", "D1"));
            Assert.Equal(new[] { "Structured Intersections!A2" }, structuredIntersection.Dependencies);
            Assert.Equal(new[] { "Structured Intersections!A2" }, structuredIntersection.FormulaDependencies);
            Assert.Equal(9, graph.EdgeCount);
            Assert.False(graph.HasCircularReferences);
        }

        [Fact]
        public void Test_FormulaDependencyGraph_ResolvesCurrentRowStructuredReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Current Row");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellFormula(2, 1, "CurrentRowData[@B]");
            sheet.CellFormula(2, 2, "CurrentRowData[[#This Row],[A]]");
            sheet.AddTable("A1:B2", hasHeader: true, name: "CurrentRowData", style: TableStyle.TableStyleMedium2);

            ExcelSheet crossSource = document.AddWorksheet("Cross Source");
            crossSource.CellFormula(2, 1, "CrossRowData[@B]");
            ExcelSheet crossTable = document.AddWorksheet("Cross Table");
            crossTable.CellValue(1, 1, "A");
            crossTable.CellValue(1, 2, "B");
            crossTable.CellValue(2, 1, 1d);
            crossTable.CellFormula(2, 2, "'Cross Source'!A2");
            crossTable.AddTable("A1:B2", hasHeader: true, name: "CrossRowData", style: TableStyle.TableStyleMedium2);

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
            ExcelSheet sheet = document.AddWorksheet("Structured Column Ranges");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "C");
            sheet.CellFormula(2, 1, "SUM(Sales[[B]:[C]])");
            sheet.CellFormula(2, 2, "A2");
            sheet.CellValue(2, 3, 1d);
            sheet.CellFormula(2, 5, "SUM(Sales[[#Data],[B]:[C]])");
            sheet.AddTable("A1:C2", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);

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
        public void Test_FormulaEvaluator_ResolvesCurrentRowStructuredReferences() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Structured Evaluation");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "At");
            sheet.CellValue(1, 4, "Explicit");
            sheet.CellValue(1, 5, "Unqualified At");
            sheet.CellValue(1, 6, "Unqualified Plain");
            sheet.CellValue(1, 7, "Unqualified Compare");
            sheet.CellValue(2, 1, 10d);
            sheet.CellValue(2, 2, 20d);
            sheet.CellFormula(2, 3, "Sales[@B]+1");
            sheet.CellFormula(2, 4, "SUM(Sales[[#This Row],[B]])");
            sheet.CellFormula(2, 5, "[@B]+1");
            sheet.CellFormula(2, 6, "SUM([B])");
            sheet.CellFormula(2, 7, "IF([@B]>10,1,0)");
            sheet.AddTable("A1:G2", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);

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
            Assert.Equal(5, calculated);
        }

        [Fact]
        public void Test_FormulaStructuredReferences_ResolveWhitespaceAndCombinedItems() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Structured Combined Items");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "C");
            sheet.CellValue(2, 1, 10d);
            sheet.CellFormula(2, 2, "1+1");
            sheet.CellValue(2, 3, 30d);
            sheet.CellValue(3, 1, 1d);
            sheet.CellValue(3, 2, 5d);
            sheet.CellValue(3, 3, 3d);
            sheet.CellFormula(1, 5, "SUM(Sales[[#Data], [A] : [C]])");
            sheet.CellFormula(2, 5, "SUM(Sales[[#Headers],[#Data],[B]])");
            sheet.CellFormula(3, 5, "SUM(Sales[[#Data],[#Totals],[B]])");
            sheet.AddTable("A1:C3", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);
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
            ExcelSheet sheet = document.AddWorksheet("Unqualified Current Row");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "Sales Amount");
            sheet.CellValue(1, 4, "% Commission");
            sheet.CellValue(1, 5, "Book.xlsx");
            sheet.CellFormula(2, 1, "[@B]");
            sheet.CellFormula(2, 2, "[[#This Row],[A]]");
            sheet.CellFormula(2, 3, "[% Commission]");
            sheet.CellFormula(2, 4, "[@[Sales Amount]]");
            sheet.CellFormula(2, 6, "[@B]");
            sheet.CellFormula(3, 1, "UnknownTable[[#This Row],[B]]");
            sheet.CellFormula(3, 2, "[@B]Sheet1!A1");
            sheet.CellValue(3, 3, 1d);
            sheet.CellValue(3, 4, 2d);
            sheet.CellFormula(3, 5, "[Book.xlsx]Sheet1!RemoteName");
            sheet.AddTable(
                "A1:E3",
                hasHeader: true,
                name: "UnqualifiedCurrentRowData",
                style: TableStyle.TableStyleMedium2);

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
            ExcelSheet sheet = document.AddWorksheet("Lexical");
            sheet.CellFormula(1, 1, "LET(Input,1,Input+1)");
            sheet.CellFormula(1, 2, "A1");
            sheet.CellFormula(2, 1, "Input+LET(Input,1,Input+1)");
            sheet.CellFormula(2, 2, "LAMBDA(Input,Input+1)(1)");
            sheet.CellFormula(3, 1, "LET(Input,{1,2},SUM(Input))");
            sheet.CellValue(1, 4, "Amount");
            sheet.CellValue(1, 5, "Other");
            sheet.CellValue(2, 4, 1d);
            sheet.CellValue(2, 5, 2d);
            sheet.AddTable("D1:E2", hasHeader: true, name: "LexicalData", style: TableStyle.TableStyleMedium2);
            sheet.CellFormula(4, 1, "LET(Input,LexicalData[[#Headers],[Amount]],Input)");
            ExcelSheet quotedQualifier = document.AddWorksheet("O'A,B)");
            quotedQualifier.CellValue(1, 1, 5d);
            sheet.CellFormula(5, 1, "LET(Input,'O''A,B)'!A1,Input)");
            sheet.CellFormula(6, 1, "LET(LexicalData,1,SUM(LexicalData[Amount]))");
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
            ExcelSheet source = document.AddWorksheet("Source");
            source.CellFormula(1, 1, "Other!C2");
            ExcelSheet other = document.AddWorksheet("Other");
            other.CellValue(1, 1, "Amount");
            other.CellValue(1, 2, "Label");
            other.CellValue(1, 3, "Result");
            other.CellValue(2, 1, 2d);
            other.CellValue(2, 2, "Row");
            other.CellFormula(2, 3, "OtherSales[@Amount]+1");
            other.AddTable("A1:C2", hasHeader: true, name: "OtherSales", style: TableStyle.TableStyleMedium2);
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
                document.SetNamedRange("Amount", "'Named Circular'!A1", save: false);
                document.SetNamedRange("SUM", "'Named Circular'!A1", save: false);

                ExcelSheet structured = document.AddWorksheet("Structured");
                structured.CellValue(1, 1, "Amount");
                structured.CellFormula(2, 1, "SUM(Sales[Amount])");
                structured.CellValue(1, 2, "A1");
                structured.CellValue(2, 2, 2d);
                structured.CellFormula(2, 3, "SUM(Sales[A1])");
                structured.AddTable("A1:B2", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);

                ExcelSheet sales = document.AddWorksheet("Sales");
                sales.CellValue(1, 1, 5d);
                ExcelSheet salesData = document.AddWorksheet("Sales Data");
                salesData.CellValue(1, 3, 7d);
                salesData.SetNamedRange("SharedInput", "C1", save: false);
                ExcelSheet apostropheSales = document.AddWorksheet("O'Sales");
                apostropheSales.CellValue(1, 3, 8d);
                apostropheSales.SetNamedRange("EscapedInput", "C1", save: false);

                ExcelSheet tokens = document.AddWorksheet("Tokens");
                tokens.CellValue(1, 1, 100d);
                tokens.CellFormula(1, 2, "SUM(A1)");
                tokens.CellFormula(1, 3, "LOG10(A1)");
                tokens.CellFormula(1, 4, "Sales!A1+1");
                tokens.CellFormula(1, 5, "'Sales Data'!C1");
                tokens.CellFormula(1, 6, "'Sales Data'!SharedInput+1");
                tokens.CellFormula(1, 7, "'O''Sales'!EscapedInput+1");
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
            ExcelSheet sheet = document.AddWorksheet("Sales, East");
            sheet.CellFormula(1, 1, "B1");
            sheet.CellFormula(1, 2, "CommaAlias");
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
            ExcelSheet columns = document.AddWorksheet("Whole Column Name");
            columns.CellFormula(1, 1, "SUM(AllB)");
            columns.CellFormula(1, 2, "A1");
            ExcelSheet rows = document.AddWorksheet("Whole Row Name");
            rows.CellFormula(1, 1, "SUM(AllRows)");
            rows.CellFormula(2, 1, "A1");
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
            ExcelSheet sheet = document.AddWorksheet("Backslash Table");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellFormula(2, 1, "C2");
            sheet.CellFormula(2, 3, "SUM(TableAlias)");
            sheet.CellFormula(1, 4, "SUM(\\Sales[Amount])");
            sheet.AddTable("A1:A2", hasHeader: true, name: "Sales", style: TableStyle.TableStyleMedium2);
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
            ExcelSheet sheet = document.AddWorksheet("Chained Names");
            sheet.CellFormula(1, 1, "B1");
            sheet.CellFormula(1, 2, "Alias");
            sheet.CellFormula(1, 3, "First");
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
            ExcelSheet sheet = document.AddWorksheet("Named Range Endpoints");
            sheet.CellValue(1, 1, 1d);
            sheet.CellFormula(1, 2, "D1");
            sheet.CellValue(1, 3, 3d);
            sheet.CellFormula(1, 4, "SUM(Start:End)");
            sheet.CellFormula(1, 5, "SUM((Start):(End))");
            sheet.CellFormula(1, 6, "SUM(Start:C1)");
            sheet.CellFormula(1, 7, "SUM(A1:End)");
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
            ExcelSheet sheet = document.AddWorksheet("Bounded Names");
            sheet.CellValue(1, 1, 1d);
            sheet.CellFormula(1, 2, "Name000");
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
        public void Test_FormulaEvaluator_DoesNotConsumeDepthBlockedValuesInInfoFunctions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFormulaDepth.InfoFunction.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Depth");
            sheet.CellFormula(1, 1, "ISBLANK(A2)");
            sheet.CellFormula(2, 1, "1+1");
            document.Calculation.MaximumDependencyDepth = 1;

            Assert.Equal(1, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 1, out _));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? cached));
            Assert.Equal("2", cached);
        }

        [Fact]
        public void Test_FormulaEvaluator_EnforcesDependencyDepthAcrossWorksheets() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet source = document.AddWorksheet("Source");
            ExcelSheet other = document.AddWorksheet("Other");
            source.CellFormula(1, 1, "Other!A1+1");
            other.CellFormula(1, 1, "1+1");
            document.Calculation.MaximumDependencyDepth = 1;

            Assert.Equal(1, document.Calculate());
            Assert.False(source.TryGetCachedFormulaValue(1, 1, out _));
            Assert.True(other.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal("2", cached);
        }

        [Fact]
        public void Test_FormulaEvaluator_DoesNotCacheAggregatesAfterDependencyDepthGuard() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Depth");
            sheet.CellFormula(1, 1, "SUBTOTAL(2,A2)");
            sheet.CellFormula(1, 2, "COUNTBLANK(B2)");
            sheet.CellFormula(1, 3, "COUNTIF(C2,\">0\")");
            sheet.CellFormula(2, 1, "1+1");
            sheet.CellFormula(2, 2, "1+1");
            sheet.CellFormula(2, 3, "1+1");
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
            ExcelSheet sheet = document.AddWorksheet("Depth");
            sheet.CellFormula(1, 1, "UNSUPPORTED()");
            sheet.CellFormula(2, 1, "SUM(A1)");
            sheet.CellFormula(3, 1, "SUM(A2)");

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
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellFormula(1, 2, "SUM(A1,$A1,A$1,$A$1)");
                sheet.CellFormula(1, 3, "SUM(B1,$A1,B$1,$A$1)");
                sheet.CellFormula(2, 2, "SUM(A2,$A2,A$1,$A$1)");
                sheet.CellFormula(2, 3, "SUM(B2,$A2,B$1,$A$1)");
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

                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                ExcelCellSnapshot snapshotFollower = Assert.Single(
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
                ExcelSheet sheet = document.AddWorksheet("Shared Formula Text");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                sheet.CellFormula(2, 2, "A2+1");
                sheet.CellFormula(1, 3, "FORMULATEXT(B2)");
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
            const string masterFormula = "SUM(A:A,$A:$A,1:1,$1:$1,A1#,$A1#,A$1#,$A$1#,A1%)+LOG10(A1)+LOG10 (A1)+FOO1 (A1)+SUM(A1 (A1:A2))+Q1:Q4!A1";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Q1");
                document.AddWorksheet("Q4");
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellFormula(1, 5, masterFormula);
                sheet.CellFormula(1, 6, masterFormula);
                sheet.CellFormula(2, 5, masterFormula);
                sheet.CellFormula(2, 6, masterFormula);
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
            loaded.Calculation.RegisterCustomFunction("FOO1", (_, _) => ExcelFormulaValue.Blank);
            ExcelSheet shared = loaded["Shared"];
            Assert.Equal(masterFormula, shared.GetFormulaText(1, 5));
            Assert.Equal(
                "SUM(B:B,$A:$A,2:2,$1:$1,B2#,$A2#,B$1#,$A$1#,B2%)+LOG10(B2)+LOG10 (B2)+FOO1 (B2)+SUM(B2 (B2:B3))+Q1:Q4!B2",
                shared.GetFormulaText(2, 6));
        }

        [Fact]
        public void Test_FormulaDependencyGraph_IgnoresDefinedNamesInsideErrorLiterals() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Errors");
            sheet.CellFormula(1, 1, "#NAME?");
            sheet.CellFormula(1, 2, "A1+1");
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
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellValue(1, 1, 10d);
            ExcelSheet first = document.AddWorksheet("First");
            first.CellValue(1, 1, 20d);
            ExcelSheet last = document.AddWorksheet("Last");
            last.CellValue(1, 1, 30d);
            document.AddWorksheet("Q1");
            document.AddWorksheet("Q4");
            ExcelSheet tokens = document.AddWorksheet("Tokens");
            tokens.CellFormula(1, 1, "[Other.xlsx]Data!A1+First:Last!A1+'First:Last'!A1+Q1:Q4!A1");
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
