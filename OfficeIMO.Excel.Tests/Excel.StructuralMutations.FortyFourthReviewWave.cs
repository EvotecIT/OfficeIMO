using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImageMutations_InvalidateFormulaAndTextCaches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.CellValue(1, 1, "Header");
            sheet.CellFormula(2, 1, "A1");
            Cell formulaCell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A2");
            formulaCell.CellValue = new CellValue("Header");
            formulaCell.DataType = CellValues.String;
            formulaCell.CellFormula!.CalculateCell = false;
            document.MarkFormulaSheetRecalculated(
                sheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());

            Assert.True(sheet.TryGetColumnIndexByHeader("Header", out int headerColumn));
            Assert.Equal(1, headerColumn);
            Assert.Equal("A1", sheet.FindFirst("Header"));
            Assert.True(Assert.Single(sheet.GetFormulaCells()).State.HasFlag(ExcelFormulaState.Evaluated));

            sheet.SetInCellImage(1, 1, TinyPng, altText: "Replacement");

            Assert.False(sheet.TryGetColumnIndexByHeader("Header", out _));
            Assert.Equal("A2", sheet.FindFirst("Header"));
            Assert.True(sheet.TryGetColumnIndexByHeader("#VALUE!", out headerColumn));
            Assert.Equal(1, headerColumn);
            Assert.Equal("A1", sheet.FindFirst("#VALUE!"));
            Assert.True(Assert.Single(sheet.GetFormulaCells()).State.HasFlag(ExcelFormulaState.Dirty));

            formulaCell.CellFormula!.CalculateCell = false;
            document.MarkFormulaSheetRecalculated(
                sheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());
            Assert.True(sheet.RemoveInCellImage(1, 1));

            Assert.False(sheet.TryGetColumnIndexByHeader("#VALUE!", out _));
            Assert.Null(sheet.FindFirst("#VALUE!"));
            ExcelFormulaCellInfo stale = Assert.Single(sheet.GetFormulaCells());
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(stale.State.HasFlag(ExcelFormulaState.Evaluated));
        }

        [Fact]
        public void Test_StructuralPlanning_BoundsUnloadedNativeConnectionMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ConnectionsPart connectionsPart = document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            byte[] payload = Encoding.UTF8.GetBytes(
                "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + new string('x', ExcelDocument.MaximumWorkbookConnectionMetadataCharacters)
                + "</connections>");
            using (var stream = new MemoryStream(payload, writable: false)) {
                connectionsPart.FeedData(stream);
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => sheet.PlanInsertRows(1));

            Assert.Contains("connection metadata exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(connectionsPart.IsRootElementLoaded);
        }

        [Fact]
        public async Task Test_QueryAndSparklineReads_WaitForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "ReadLockQuery",
                WorksheetName = sheet.Name,
                TableName = "ReadLockResults",
                ColumnNames = new[] { "Value" }
            });
            sheet.AddSparklines("A1:B1", "C1");

            IReadOnlyList<ExcelQueryBackedTableInfo> queries = await AssertWorkbookReadWaitsForWriter(
                document,
                document.GetQueryBackedTables,
                "query-backed table");
            IReadOnlyList<ExcelSparklineInfo> sparklines = await AssertWorkbookReadWaitsForWriter(
                document,
                sheet.GetSparklines,
                "sparkline");

            Assert.Single(queries);
            Assert.Single(sparklines);
        }

        [Fact]
        public void Test_FormulaSyntaxTree_KeepsDeletedSheetAddressesOpaque() {
            const string formula = "=#REF!A1+#REF!A1:C3+SUM(B2)";

            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            ExcelFormulaReferenceSyntax reference = Assert.Single(tree.Nodes.OfType<ExcelFormulaReferenceSyntax>());
            Assert.Equal("B2", reference.Text);
            Assert.Equal(
                "=#REF!A1+#REF!A1:C3+SUM(C2)",
                tree.Rewrite(item => item.Offset(0, 1)));
        }

        private static async Task<IReadOnlyList<T>> AssertWorkbookReadWaitsForWriter<T>(
            ExcelDocument document,
            Func<IReadOnlyList<T>> readAction,
            string description) {
            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterWriteLock();
            Task<IReadOnlyList<T>>? read = null;
            using var started = new ManualResetEventSlim(false);
            try {
                read = Task.Factory.StartNew(
                    () => {
                        started.Set();
                        return readAction();
                    },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingReadCount > 0 || read.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"The {description} read did not reach the workbook lock (status: {read.Status}).");
                Assert.False(read.IsCompleted);
            } finally {
                workbookLock.ExitWriteLock();
            }

            return await read!;
        }
    }
}
