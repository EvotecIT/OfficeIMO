using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaAuthoring_PreservesRetainedCachesButMarksEveryFormulaKindDirty() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");

            sheet.CellValue(1, 1, 4d);
            sheet.CellValue(1, 2, 4d);
            sheet.CellValue(1, 3, 4d);
            sheet.CellFormula(1, 1, "2+2");
            sheet.SetArrayFormula("B1:B1", "2+2");
            sheet.SetLegacyArrayFormula("C1:C1", "2+2");

            Assert.All(sheet.GetFormulaCells(), formula => {
                Assert.Equal("4", formula.CachedValue);
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Dirty));
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Deferred));
                Assert.False(formula.State.HasFlag(ExcelFormulaState.Evaluated));
            });

            sheet.CellFormula(1, 1, "3+3");

            Assert.All(sheet.GetFormulaCells(), formula => {
                Assert.Equal("4", formula.CachedValue);
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Dirty));
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Deferred));
                Assert.False(formula.State.HasFlag(ExcelFormulaState.Evaluated));
            });
        }

        [Fact]
        public void Test_NamedStyleValidation_RejectsInvalidXmlBeforeStylesheetMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetBold();
            string originalStyles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.OuterXml;

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                sheet.DefineNamedStyle("Bad\u0001Style", 1, 1));

            Assert.Equal("name", exception.ParamName);
            Assert.Equal(originalStyles, document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.OuterXml);
            Assert.DoesNotContain(document.GetNamedStyles(), style => style.Name.Contains("Bad", StringComparison.Ordinal));
        }

        [Fact]
        public async Task Test_QueryBackedTable_RejectsDuplicateConnectionInsideTransactionalOwner() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            using var start = new ManualResetEventSlim(false);
            int successes = 0;
            Task[] callers = Enumerable.Range(0, 8).Select(index => Task.Run(() => {
                start.Wait();
                try {
                    document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                        ConnectionName = "SharedConnection",
                        WorksheetName = sheet.Name,
                        StartCell = $"A{index + 1}",
                        TableName = $"Results{index + 1}",
                        ColumnNames = new[] { "Value" }
                    });
                    Interlocked.Increment(ref successes);
                } catch (InvalidOperationException) {
                    // Exactly one caller owns the connection; every contender must be rejected.
                }
            })).ToArray();

            start.Set();
            await Task.WhenAll(callers);

            Assert.Equal(1, successes);
            ExcelQueryBackedTableInfo query = Assert.Single(document.GetQueryBackedTables());
            Assert.StartsWith("Results", query.TableName, StringComparison.Ordinal);
            Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public async Task Test_SheetBatch_DoesNotExposeItsUnlockedFastPathToAnotherThread() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            using var batchEntered = new ManualResetEventSlim(false);
            using var releaseBatch = new ManualResetEventSlim(false);

            Task batch = Task.Factory.StartNew(
                () => sheet.Batch(_ => {
                    batchEntered.Set();
                    releaseBatch.Wait();
                }),
                CancellationToken.None,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
            Assert.True(batchEntered.Wait(TimeSpan.FromSeconds(10)));

            Task write = Task.Factory.StartNew(
                () => sheet.CellValue(1, 1, "Serialized"),
                CancellationToken.None,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
            try {
                Assert.True(
                    SpinWait.SpinUntil(
                        () => workbookLock.WaitingWriteCount > 0 || write.IsCompleted,
                        TimeSpan.FromSeconds(10)),
                    $"Cell write did not reach the workbook lock (status: {write.Status}).");
                Assert.False(write.IsCompleted);
            } finally {
                releaseBatch.Set();
            }

            await batch;
            await write;
            Assert.Equal("Serialized", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
