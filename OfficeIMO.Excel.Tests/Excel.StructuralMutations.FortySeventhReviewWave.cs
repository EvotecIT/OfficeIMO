using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_RemoveQueryBackedTable_RevalidatesConnectionInsideTransaction() {
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

            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterReadLock();
            Task<bool>? removal = null;
            using var started = new ManualResetEventSlim(false);
            ExcelQueryBackedTableInfo? replacement = null;
            try {
                removal = Task.Factory.StartNew(
                    () => {
                        started.Set();
                        return document.RemoveQueryBackedTable(original.TableName, preserveTable: false);
                    },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingWriteCount > 0 || removal.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"Query removal did not reach the transaction boundary (status: {removal.Status}).");
                Assert.False(removal.IsCompleted);
                using (data.BeginNoLock()) {
                    Assert.True(document.RemoveQueryBackedTable(original.TableName, preserveTable: false));
                    replacement = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                        ConnectionName = "ReplacementQuery",
                        WorksheetName = data.Name,
                        TableName = original.TableName,
                        ColumnNames = new[] { "Value" }
                    });
                }
            } finally {
                workbookLock.ExitReadLock();
            }

            await Assert.ThrowsAsync<InvalidOperationException>(async () => await removal!);
            Assert.NotNull(replacement);
            Assert.NotEqual(original.ConnectionId, replacement!.ConnectionId);
            Assert.Equal(replacement.ConnectionId, document.GetQueryBackedTables()
                .Single(item => item.TableName == replacement.TableName).ConnectionId);
        }

        [Fact]
        public void Test_AddQueryBackedTable_RejectsOrdinaryConnectionNameConflict() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ConnectionsPart connectionsPart = document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            connectionsPart.Connections = new Connections(new Connection {
                Id = 42U,
                Name = "ExistingConnection",
                Type = 1U
            });
            connectionsPart.Connections.Save();

            Assert.Throws<InvalidOperationException>(() => document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "existingconnection",
                WorksheetName = sheet.Name,
                TableName = "Results",
                ColumnNames = new[] { "Value" }
            }));

            Assert.Empty(sheet.WorksheetPart.TableDefinitionParts);
            Assert.Null(sheet.CellAt(1, 1).GetValue<string>());
            Assert.Single(connectionsPart.Connections.Elements<Connection>());
        }

        [Fact]
        public async Task Test_AddPivotTimeline_RevalidatesPivotInsideTransaction() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellValue(1, 1, "OrderDate");
            data.CellValue(1, 2, "Sales");
            data.CellValue(2, 1, new DateTime(2026, 1, 2));
            data.CellValue(2, 2, 10d);
            data.AddPivotTable(
                "A1:B2",
                "D2",
                "SalesPivot",
                rowFields: new[] { "OrderDate" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
            _ = document.Sheets;

            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterReadLock();
            Task<ExcelPivotInteractionInfo>? creation = null;
            using var started = new ManualResetEventSlim(false);
            try {
                creation = Task.Factory.StartNew(
                    () => {
                        started.Set();
                        return document.AddPivotTimeline("SalesPivot", "OrderDate", data.Name);
                    },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                Assert.True(started.Wait(TimeSpan.FromSeconds(10)));
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingWriteCount > 0 || creation.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"Timeline creation did not reach the transaction boundary (status: {creation.Status}).");
                Assert.False(creation.IsCompleted);
                using (data.BeginNoLock()) {
                    data.WorksheetPart.DeletePart(Assert.Single(data.WorksheetPart.PivotTableParts));
                }
            } finally {
                workbookLock.ExitReadLock();
            }

            await Assert.ThrowsAsync<ArgumentException>(async () => await creation!);
            Assert.Empty(document.GetPivotInteractions());
            Assert.Empty(document.WorkbookPartRoot.TimeLineCacheParts);
            Assert.Empty(data.WorksheetPart.TimeLineParts);
        }

        [Fact]
        public void Test_RenameTable_SynchronizesAttachedQueryTableName() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "RenameQuery",
                WorksheetName = sheet.Name,
                TableName = "OriginalResults",
                ColumnNames = new[] { "Value" }
            });
            TableDefinitionPart tablePart = Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            QueryTablePart queryPart = Assert.Single(tablePart.QueryTableParts);
            Assert.Equal(source.TableName, queryPart.QueryTable!.Name!.Value);

            Assert.Equal("RenamedResults", sheet.RenameTable(source.TableName, "RenamedResults"));

            Assert.Equal("RenamedResults", tablePart.Table!.Name!.Value);
            Assert.Equal("RenamedResults", tablePart.Table.DisplayName!.Value);
            Assert.Equal("RenamedResults", queryPart.QueryTable!.Name!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernChart_RenameRejectsDuplicateWorksheetDrawingName() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            var data = new ExcelChartData(
                new[] { "A" },
                new[] { new ExcelChartSeries("Value", new[] { 1d }) });
            ExcelModernChart first = sheet.AddModernChart(data, 1, 1, ExcelModernChartType.Funnel);
            ExcelModernChart second = sheet.AddModernChart(data, 20, 1, ExcelModernChartType.Treemap);
            first.Name = "SharedName";
            string secondName = second.Name;

            Assert.Throws<InvalidOperationException>(() => second.Name = "sharedname");

            Assert.Equal(secondName, second.Name);
            Assert.Equal(first.Name, sheet.GetModernChart("SharedName")!.Name);
            Assert.Equal(2, sheet.ModernCharts.Count());
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
