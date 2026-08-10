using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImage_MetadataBudgetStopsBeforeUnboundedRootMaterialization() {
            using var source = new CountingNonSeekableReadStream(length: 1_000_000);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelSheet.ValidateInCellImageMetadataStream(source, 32, "Rich-value data"));

            Assert.Contains("32-byte limit", exception.Message, StringComparison.Ordinal);
            Assert.InRange(source.BytesRead, 33, 100_000);

            using var seekable = new MemoryStream(new byte[33]);
            seekable.Position = 7;
            Assert.Throws<InvalidDataException>(() =>
                ExcelSheet.ValidateInCellImageMetadataStream(seekable, 32, "Cell metadata"));
            Assert.Equal(7, seekable.Position);
        }

        [Fact]
        public async Task Test_AllowedEditRangeProtectionCheckIsAtomicWithUnprotect() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");

            for (int iteration = 0; iteration < 32; iteration++) {
                sheet.Protect();
                using var start = new Barrier(2);
                Task setter = Task.Factory.StartNew(() => {
                    start.SignalAndWait();
                    try {
                        sheet.SetAllowedEditRange("Inputs", new[] { "A1" });
                    } catch (InvalidOperationException) {
                    }
                }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
                Task unprotect = Task.Factory.StartNew(() => {
                    start.SignalAndWait();
                    sheet.Unprotect();
                }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

                await Task.WhenAll(setter, unprotect);

                Assert.False(sheet.IsProtected);
                Assert.Empty(sheet.WorksheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges>());
            }
        }

        [Theory]
        [InlineData("Sheet1!A1:Sheet1!C1", "Sheet1!A1:C1")]
        [InlineData("'Data Set'!$A$1:'Data Set'!$C$1", "'Data Set'!$A$1:$C$1")]
        [InlineData("Data!2:Data!5", "Data!2:5")]
        [InlineData("Data!A:Data!C", "Data!A:C")]
        [InlineData("'Data: Q1'!A1:'Data: Q1'!C3", "'Data: Q1'!A1:C3")]
        [InlineData("'Owner''s Data'!A1:'Owner''s Data'!C3", "'Owner''s Data'!A1:C3")]
        public void Test_FormulaSyntaxTree_NormalizesRepeatedIdenticalRangeQualifiers(
            string formula,
            string normalized) {
            Assert.True(ExcelFormulaReferenceRewriter.TryReadReferenceAt(formula, 0, out ExcelFormulaReferenceCandidate? candidate));
            Assert.NotNull(candidate);
            Assert.Equal(formula, candidate!.Text);
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            ExcelFormulaReferenceSyntax reference = Assert.Single(tree.Nodes.OfType<ExcelFormulaReferenceSyntax>());
            Assert.Equal(formula, reference.Text);
            Assert.Equal(normalized, tree.Rewrite(item => item));
        }

        [Fact]
        public void Test_StructuralColumns_RewriteRepeatedIdenticalRangeQualifiersAsOneRange() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet formulas = document.AddWorksheet("Formulas");
            formulas.CellFormula(1, 1, "SUM(Data!A1:Data!C1)");

            data.DeleteColumns(1);

            Assert.Equal("SUM(Data!A1:B1)", Assert.Single(formulas.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_FormulaSyntaxTree_IdentifiesNestedFunctionsWithoutInspectingStringsOrStructuredColumns() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(
                "IF(A1>0, _xlfn.XLOOKUP(A1,Data[Call (legacy)],Data[Value]),\"SUM(B1)\")");

            Assert.Equal(
                new[] { "IF", "_xlfn.XLOOKUP" },
                tree.Nodes.OfType<ExcelFormulaFunctionSyntax>().Select(function => function.Name));
        }

        [Fact]
        public void Test_FormulaSyntaxTree_RewritesOnlyCompleteDefinedNameNodes() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(
                "SUM(Tax,TaxRate,\"Tax\")+Table1[Tax]+Sheet1!Tax");

            string rewritten = tree.RewriteNames(name =>
                string.Equals(name, "Tax", StringComparison.OrdinalIgnoreCase) ? "Tax_Copy" : name);

            Assert.Equal("SUM(Tax_Copy,TaxRate,\"Tax\")+Table1[Tax]+Sheet1!Tax", rewritten);
        }

        [Fact]
        public void Test_FormulaExpressionParser_ParsesCustomFunctionCallWithoutClassifyingItAsBuiltIn() {
            Assert.True(ExcelFormulaExpressionParser.TryParseFunctionCall(
                "DOUBLEVALUE(A1)", out ExcelFormulaFunctionCallSyntax? call));
            Assert.Equal("DOUBLEVALUE", call!.Name);
            Assert.Equal("A1", call.Arguments);
            Assert.False(ExcelFormulaExpressionParser.TryParseSupportedFunctionCall("DOUBLEVALUE(A1)", out _));
        }

        [Theory]
        [InlineData("=SUM('Q1) final'!A1:A2)", "'Q1) final'!A1:A2")]
        [InlineData("=SUM('Owner''s (Q1)'!A1:A2)", "'Owner''s (Q1)'!A1:A2")]
        public void Test_FormulaExpressionParser_IgnoresParenthesesInsideQuotedSheetQualifiers(
            string formula,
            string expectedArguments) {
            Assert.True(ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(
                formula, out ExcelFormulaFunctionCallSyntax? call));
            Assert.Equal("SUM", call!.Name);
            Assert.Equal(expectedArguments, call.Arguments);
        }

        [Fact]
        public void Test_FormulaExpressionParser_RejectsUnterminatedQualifierInsideFunction() {
            Assert.False(ExcelFormulaExpressionParser.TryParseFunctionCall("=SUM('Q1) final!A1:A2)", out _));
        }

        [Fact]
        public void Test_FormulaExpressionParser_TreatsApostrophesInsideStructuredReferencesAsEscapes() {
            Assert.True(ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(
                "=SUM(Table1['#Data])", out ExcelFormulaFunctionCallSyntax? call));
            Assert.Equal("Table1['#Data]", call!.Arguments);

            Assert.True(ExcelFormulaExpressionParser.TryParseArithmetic(
                "Table1['#Data]+1", out ExcelFormulaBinaryExpressionSyntax? expression));
            Assert.Equal("Table1['#Data]", expression!.Left);
            Assert.Equal("+", expression.Operator);
            Assert.Equal("1", expression.Right);
        }

        [Theory]
        [InlineData("'Profit-Loss'!A1+1", "'Profit-Loss'!A1", "+", "1")]
        [InlineData("'Owner''s <Data>'!A1>=0", "'Owner''s <Data>'!A1", ">=", "0")]
        public void Test_FormulaExpressionParser_IgnoresOperatorsInsideQuotedSheetQualifiers(
            string formula,
            string expectedLeft,
            string expectedOperator,
            string expectedRight) {
            ExcelFormulaBinaryExpressionSyntax? expression;
            bool parsed = expectedOperator == "+"
                ? ExcelFormulaExpressionParser.TryParseArithmetic(formula, out expression)
                : ExcelFormulaExpressionParser.TryParseComparison(formula, out expression);

            Assert.True(parsed);
            Assert.NotNull(expression);
            Assert.Equal(expectedLeft, expression!.Left);
            Assert.Equal(expectedOperator, expression.Operator);
            Assert.Equal(expectedRight, expression.Right);
        }

        [Theory]
        [InlineData("'Unclosed-Sheet!A1+1")]
        [InlineData("'Unclosed<Sheet!A1=1")]
        public void Test_FormulaExpressionParser_RejectsUnterminatedQuotedSheetQualifiers(string formula) {
            Assert.False(ExcelFormulaExpressionParser.TryParseArithmetic(formula, out _));
            Assert.False(ExcelFormulaExpressionParser.TryParseComparison(formula, out _));
        }

        [Fact]
        public void Test_NamedRangeParser_PreservesBangInsideQuotedSheetName() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Bang!Sheet");

            document.SetNamedRange("DataRange", "'Bang!Sheet'!A1:B2", save: false,
                validationMode: ExcelDefinedNameValidationMode.Strict);

            DocumentFormat.OpenXml.Spreadsheet.DefinedName definedName = Assert.Single(
                document.WorkbookPartRoot.Workbook.DefinedNames!
                    .Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>());
            Assert.Equal("'Bang!Sheet'!$A$1:$B$2", definedName.Text);
        }

        [Fact]
        public void Test_SparklineType_RejectsUndefinedValueBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AddSparklines("A1:C1", "D1");
            string originalXml = sheet.WorksheetPart.Worksheet.OuterXml;

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                sheet.SetSparklineType("D1", (ExcelSparklineType)int.MaxValue));

            Assert.Equal(originalXml, sheet.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(ExcelSparklineType.Line, Assert.Single(sheet.GetSparklines()).Type);
        }

        [Fact]
        public void Test_ModernChartDataRange_RejectsXfdGrowthBeforeUpdateCanWrite() {
            var range = new ExcelChartDataRange(
                "Data",
                startRow: 1,
                startColumn: A1.MaxColumns - 1,
                categoryCount: 1,
                seriesCount: 1);

            Assert.Throws<ArgumentOutOfRangeException>(() => range.WithSize(1, 2));
            Assert.Equal(A1.MaxColumns, range.SeriesEndColumn);
            Assert.Equal(1, range.SeriesCount);
        }
    }
}