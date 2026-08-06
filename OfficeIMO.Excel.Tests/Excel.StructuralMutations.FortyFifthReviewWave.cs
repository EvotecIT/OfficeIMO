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

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaAuthoring_InvalidatesPreviouslyEvaluatedDependents() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");
            ExcelSheet crossSheet = document.AddWorksheet("CrossSheet");
            sheet.CellFormula(2, 1, "A1+1");
            crossSheet.CellFormula(1, 1, "Calc!A1+1");
            Cell dependent = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A2");
            dependent.CellValue = new CellValue("1");
            dependent.DataType = CellValues.Number;
            dependent.CellFormula!.CalculateCell = false;
            Cell crossSheetDependent = crossSheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            crossSheetDependent.CellValue = new CellValue("1");
            crossSheetDependent.DataType = CellValues.Number;
            crossSheetDependent.CellFormula!.CalculateCell = false;
            document.MarkFormulaSheetRecalculated(
                sheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());
            document.MarkFormulaSheetRecalculated(
                crossSheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());

            Assert.All(sheet.GetFormulaCells().Concat(crossSheet.GetFormulaCells()), baseline => {
                Assert.True(baseline.State.HasFlag(ExcelFormulaState.Cached));
                Assert.False(baseline.State.HasFlag(ExcelFormulaState.Dirty));
            });

            sheet.CellFormula(1, 1, "2+2");

            ExcelFormulaCellInfo authored = sheet.GetFormulaCells().Single(item => item.CellReference == "A1");
            ExcelFormulaCellInfo stale = sheet.GetFormulaCells().Single(item => item.CellReference == "A2");
            Assert.True(authored.State.HasFlag(ExcelFormulaState.Authored));
            Assert.True(authored.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(authored.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(stale.State.HasFlag(ExcelFormulaState.Evaluated));
            ExcelFormulaCellInfo crossSheetStale = Assert.Single(crossSheet.GetFormulaCells());
            Assert.True(crossSheetStale.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(crossSheetStale.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(crossSheetStale.State.HasFlag(ExcelFormulaState.Evaluated));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_ArrayFormulaAuthoring_ClearsOwnerAndSpillInCellImages(bool legacy) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Owner");
            sheet.SetInCellImage(1, 2, TinyPng, altText: "Spill");

            if (legacy) sheet.SetLegacyArrayFormula("A1:B1", "{1,2}");
            else sheet.SetArrayFormula("A1:B1", "{1,2}");

            Assert.Empty(sheet.GetInCellImages());
            Assert.All(
                sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value is "A1" or "B1"),
                cell => Assert.Null(cell.ValueMetaIndex));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaInspection_BoundsUnloadedCellMetadataBeforeMaterialization() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");
            sheet.CellFormula(1, 1, "1+1");
            CellMetadataPart metadataPart = document.WorkbookPartRoot.AddNewPart<CellMetadataPart>();
            using (var payload = new MemoryStream(new byte[16 * 1024 * 1024 + 1], writable: false)) {
                metadataPart.FeedData(payload);
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(sheet.GetFormulaCells);

            Assert.Contains("Cell metadata exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(metadataPart.IsRootElementLoaded);
        }

        [Fact]
        public void Test_FileBackedValidationOwner_ObservesCancellationToken() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedValidationCancellation.xlsx");
            using (var document = ExcelDocument.Create()) {
                document.AddWorksheet("Data").CellValue(1, 1, "value");
                document.Save(path);
            }
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocument.ThrowIfOpenXmlValidationFails(
                    path,
                    new ExcelSaveOptions { ValidateOpenXml = true },
                    cancellation.Token));
        }

        [Theory]
        [InlineData(false, "A1:B3")]
        [InlineData(true, "A1:B2")]
        public void Test_AutoFilterCriteria_UsesTableOwnershipWithoutExistingFilter(
            bool totalsRow,
            string expectedFilterRange) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "A");
            sheet.CellValue(2, 2, 1);
            sheet.CellValue(3, 1, "B");
            sheet.CellValue(3, 2, 2);
            sheet.AddTable(
                "A1:B3",
                hasHeader: true,
                name: "Results",
                style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2,
                includeAutoFilter: false);
            if (totalsRow) {
                sheet.SetTableTotalsByName("Results", new Dictionary<string, ExcelTableTotalsFunction> {
                    ["Value"] = ExcelTableTotalsFunction.Sum
                });
            }

            sheet.ApplyAutoFilterBlankCriteria("A1:B3", 0U);

            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            Assert.Equal(expectedFilterRange, table.AutoFilter?.Reference?.Value);
            Assert.Single(table.AutoFilter!.Elements<FilterColumn>());
            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public async Task Test_ModernChartMutations_WaitForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            ReaderWriterLockSlim workbookLock = document.EnsureLock();
            workbookLock.EnterWriteLock();
            Task<ExcelModernChart>? mutation = null;
            try {
                mutation = Task.Run(() => chart.SetTitle("Blocked title"));
                Assert.True(SpinWait.SpinUntil(
                    () => workbookLock.WaitingWriteCount > 0 || mutation.IsCompleted,
                    TimeSpan.FromSeconds(10)));
                Assert.False(mutation.IsCompleted);
            } finally {
                workbookLock.ExitWriteLock();
            }

            Assert.Same(chart, await mutation!);
            Assert.Equal("Blocked title", chart.Title);
        }

        [Fact]
        public async Task Test_PivotInteractionSnapshot_WaitsForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Data");

            IReadOnlyList<ExcelPivotInteractionInfo> interactions = await AssertWorkbookReadWaitsForWriter(
                document,
                document.GetPivotInteractions,
                "pivot interaction");

            Assert.Empty(interactions);
        }
    }
}
