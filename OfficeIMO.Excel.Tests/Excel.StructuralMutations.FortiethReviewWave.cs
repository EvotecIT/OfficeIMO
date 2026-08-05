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
        public void Test_FormulaSyntaxTree_NormalizesRepeatedIdenticalRangeQualifiers(
            string formula,
            string normalized) {
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
