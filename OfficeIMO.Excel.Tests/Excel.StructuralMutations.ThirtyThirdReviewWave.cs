using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernChart_RejectsSeriesThatWouldOverflowTheBackingSheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            var series = Enumerable.Repeat(
                new ExcelChartSeries("Value", new[] { 1d }),
                A1.MaxColumns).ToArray();
            var data = new ExcelChartData(new[] { "A" }, series);

            ArgumentException exception = Assert.Throws<ArgumentException>(() => sheet.AddModernChart(
                data,
                1,
                1,
                ExcelModernChartType.Funnel));

            Assert.Equal("data", exception.ParamName);
            Assert.Contains("column limit", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(sheet.ModernCharts);
            Assert.Single(document.Sheets);
        }

        [Fact]
        public void Test_QueryBackedTable_RequiresResolvableConnectionMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "MissingConnection",
                WorksheetName = sheet.Name,
                TableName = "QueryResults",
                ColumnNames = new[] { "Value" }
            });
            Connections connections = document.WorkbookPartRoot.ConnectionsPart!.Connections!;
            Assert.Single(connections.Elements<Connection>()).Remove();
            connections.Save();

            Assert.Empty(document.GetQueryBackedTables());
            ExcelFeatureFinding finding = Assert.Single(document.InspectFeatures()
                .FindFeatures("Connections and query tables"));
            Assert.Contains(finding.Details, detail =>
                detail.Contains("queryTable", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_ModernChart_UpdateDataPreservesSeriesFormattingAndIdentity() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Original", new[] { 1d, 2d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            ExtendedChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ExtendedChartParts);
            Cx.Series originalSeries = Assert.Single(part.ChartSpace!.Descendants<Cx.Series>());
            string uniqueId = originalSeries.GetAttribute("uniqueId", string.Empty).Value;
            var formatting = new Cx.ShapeProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }));
            originalSeries.InsertAfter(formatting, originalSeries.GetFirstChild<Cx.Text>());

            chart.UpdateData(new ExcelChartData(
                new[] { "A", "B" },
                new[] { new ExcelChartSeries("Updated", new[] { 3d, 4d }) }));

            Cx.Series updatedSeries = Assert.Single(part.ChartSpace.Descendants<Cx.Series>());
            Assert.Same(originalSeries, updatedSeries);
            Assert.Same(formatting, updatedSeries.GetFirstChild<Cx.ShapeProperties>());
            Assert.Equal("FF0000", updatedSeries.GetFirstChild<Cx.ShapeProperties>()!
                .GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal(uniqueId, updatedSeries.GetAttribute("uniqueId", string.Empty).Value);
            Assert.Equal("Updated", updatedSeries.Descendants<Cx.VXsdstring>().First().Text);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_TableResize_MapsSortReferencesToDataBoundaryAboveTotalsRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.AddTable("A1:B4", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            var sortState = new SortState(
                new SortCondition { Reference = "A2:A3" }) {
                Reference = "A1:B3"
            };
            table.InsertBefore(sortState, table.TableColumns);

            sheet.ResizeTable("Sales", "A1:B6");

            Assert.Equal("A1:B5", sortState.Reference!.Value);
            Assert.Equal("A2:A5", Assert.Single(sortState.Elements<SortCondition>()).Reference!.Value);
            Assert.Equal("A1:B5", table.GetFirstChild<AutoFilter>()!.Reference!.Value);

            sheet.ResizeTable("Sales", "A1:B4");

            Assert.Equal("A1:B3", sortState.Reference!.Value);
            Assert.Equal("A2:A3", Assert.Single(sortState.Elements<SortCondition>()).Reference!.Value);
            Assert.Equal("A1:B3", table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
