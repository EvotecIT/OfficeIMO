using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesWorksheetRangeMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Worksheet worksheet = sheet.WorksheetPart.Worksheet;
            sheet.AutoFilterAdd("A5:B6");
            worksheet.Append(
                new IgnoredErrors(
                    new IgnoredError {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "C5:C6" },
                        NumberStoredAsText = true
                    }),
                new Scenarios(
                    new Scenario(
                        new InputCells { CellReference = "D5" }) {
                        Name = "Plan",
                        Count = 1U
                    }),
                new CellWatches(
                    new CellWatch { CellReference = "E5" }),
                new RowBreaks(
                    new Break {
                        Id = 5U,
                        Min = 0U,
                        Max = 16_383U,
                        ManualPageBreak = true
                    }) {
                    Count = 1U,
                    ManualBreakCount = 1U
                });
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
            worksheet.InsertBefore(
                new SheetViews(
                    new SheetView(
                        new Selection {
                            ActiveCell = "F5",
                            SequenceOfReferences = new ListValue<StringValue> {
                                InnerText = "F5:F6"
                            }
                        }) {
                        WorkbookViewId = 0U
                    }),
                sheetData);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact metadata = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "worksheet-range-metadata");
            Assert.True(metadata.ItemCount >= 6);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesQueryTableSortRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddQueryTableRefreshSortState(document, sheet, "A5:C10");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact sorts = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "query-table-sorts");
            Assert.Equal(1, sorts.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_QueryTableSortInspectionUsesChargedSnapshot() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddQueryTableRefreshSortState(document, sheet, "A5:C10");

            ExcelRowMutationPlan baseline = sheet.PlanInsertRows(5);
            ExcelRowMutationPlan bounded = sheet.PlanInsertRows(
                5,
                options: new ExcelMutationPlanOptions {
                    MaximumScannedElements = baseline.ScannedElements
                });

            Assert.Equal(baseline.ScannedElements, bounded.ScannedElements);
            Assert.Contains(
                bounded.Impacts,
                impact => impact.Category == "query-table-sorts"
                    && impact.ItemCount == 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesChartCacheInvalidation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = CreateChartOwner(document, "Summary");
            ChartPart chartPart = Assert.Single(summary.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = chartPart.ChartSpace.Descendants<C.Formula>().First();
            formula.Text = "Summary!A1:A2";
            OpenXmlElement cache = formula.Parent is C.NumberReference
                ? new C.NumberingCache()
                : new C.StringCache();
            formula.Parent!.Append(cache);

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "drawings" && impact.ItemCount >= 1);
            Assert.NotNull(cache.Parent);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanDoesNotRegisterTransientSheetWrappers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            document.AddWorksheet("Summary");
            document.AddWorksheet("Archive");
            int registeredBefore = document.RegisteredSheetWrapperCountForTests;

            for (int iteration = 0; iteration < 5; iteration++) {
                data.PlanInsertRows(5);
            }

            Assert.Equal(
                registeredBefore,
                document.RegisteredSheetWrapperCountForTests);
        }
    }
}
