using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_QueryRefresh_RejectsRecreatedBindingBeforeWritingResults() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo original = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "OriginalQuery",
                WorksheetName = data.Name,
                TableName = "Results",
                ColumnNames = new[] { "Value" }
            });
            ExcelSheet anchor = document.AddWorksheet("Anchor");
            document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "AnchorQuery",
                WorksheetName = anchor.Name,
                TableName = "AnchorResults",
                ColumnNames = new[] { "Value" }
            });
            var host = new BlockingQueryHost(new ExcelQueryExecutionResult(
                new[] { "Value" },
                new IReadOnlyList<object?>[] { new object?[] { "stale" } }));

            Task<ExcelQueryRefreshResult> refresh = document.RefreshQueryAsync(
                original.TableName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true });
            await host.Started;
            Assert.True(document.RemoveQueryBackedTable(original.TableName, preserveTable: false));
            ExcelQueryBackedTableInfo replacement = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "ReplacementQuery",
                WorksheetName = data.Name,
                TableName = original.TableName,
                ColumnNames = new[] { "Value" }
            });
            Assert.NotEqual(original.ConnectionId, replacement.ConnectionId);
            host.Release();

            await Assert.ThrowsAsync<InvalidOperationException>(() => refresh);
            Assert.Equal("Value", data.CellAt(1, 1).GetValue<string>());
            Assert.Null(data.CellAt(2, 1).GetValue<string>());
            Assert.Equal(replacement.ConnectionId, document.GetQueryBackedTables()
                .Single(item => item.TableName == replacement.TableName).ConnectionId);
        }

        [Fact]
        public async Task Test_PivotInteraction_ExplicitNameIsUniqueInsideTransaction() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellValue(1, 1, "Region");
            data.CellValue(1, 2, "Sales");
            data.CellValue(2, 1, "East");
            data.CellValue(2, 2, 10d);
            data.AddPivotTable(
                "A1:B2",
                "D2",
                "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });
            using var start = new ManualResetEventSlim(false);
            int successes = 0;
            Task[] callers = Enumerable.Range(0, 8).Select(_ => Task.Run(() => {
                start.Wait();
                try {
                    document.AddPivotSlicer(
                        "SalesPivot",
                        "Region",
                        data.Name,
                        new ExcelSlicerViewOptions { Name = "SharedFilter", Row = 6, Column = 1 });
                    Interlocked.Increment(ref successes);
                } catch (InvalidOperationException) {
                    // The first transaction owns the explicit name; every contender must be rejected.
                }
            })).ToArray();

            start.Set();
            await Task.WhenAll(callers);

            Assert.Equal(1, successes);
            Assert.Equal("SharedFilter", Assert.Single(document.GetPivotInteractions()).Name);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernChart_ValidatesPlotAreaBeforeWritingBackingData() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            ExcelChartDataRange range = chart.DataRange!;
            ExcelSheet backingSheet = document[range.SheetName];
            double originalValue = backingSheet
                .CellAt(range.SeriesStartRow, range.SeriesStartColumn)
                .GetValue<double>();
            ExtendedChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ExtendedChartParts);
            part.ChartSpace!.Descendants<Cx.PlotArea>().Single().Remove();
            part.ChartSpace.Save();

            Assert.Throws<InvalidOperationException>(() => chart.UpdateData(new ExcelChartData(
                new[] { "A" },
                new[] { new ExcelChartSeries("Value", new[] { 9d }) })));

            Assert.Equal(originalValue, backingSheet
                .CellAt(range.SeriesStartRow, range.SeriesStartColumn)
                .GetValue<double>());
        }

        [Fact]
        public async Task Test_GetNamedStyles_WaitsForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetBold();
            sheet.DefineNamedStyle("LockedStyle", 1, 1);
            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterWriteLock();
            Task<IReadOnlyList<ExcelNamedStyleInfo>>? read = null;
            using var started = new ManualResetEventSlim(false);
            try {
                read = Task.Factory.StartNew(
                    () => {
                        started.Set();
                        return document.GetNamedStyles();
                    },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingReadCount > 0 || read.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"Named-style read did not reach the workbook lock (status: {read.Status}).");
                Assert.False(read.IsCompleted);
            } finally {
                workbookLock.ExitWriteLock();
            }

            Assert.NotEmpty(await read!);
        }

        private sealed class BlockingQueryHost : IExcelQueryExecutionHost {
            private readonly TaskCompletionSource<bool> _release =
                new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly ExcelQueryExecutionResult _result;
            private readonly TaskCompletionSource<bool> _started =
                new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            internal BlockingQueryHost(ExcelQueryExecutionResult result) {
                _result = result;
            }

            internal Task Started => _started.Task;

            internal void Release() => _release.TrySetResult(true);

            public async Task<ExcelQueryExecutionResult> ExecuteAsync(
                ExcelQueryExecutionRequest request,
                CancellationToken cancellationToken) {
                _started.TrySetResult(true);
                await _release.Task.ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                return _result;
            }
        }
    }
}
