using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralInsertion_RejectsVmlNoteEndpointsAtWorksheetLimits() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Columns");
                sheet.SetComment(1, 2, "Keep", author: "Tester");
                VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
                SetCommentVmlAnchor(vmlPart, "16382, 15, 0, 2, 16383, 15, 3, 4");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertColumns(A1.MaxColumns - 1));

                Assert.Contains("comment note anchor", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.True(sheet.HasComment(1, 2));
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Rows");
                sheet.SetComment(1, 2, "Keep", author: "Tester");
                VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
                SetCommentVmlAnchor(vmlPart, $"1, 15, {A1.MaxRows - 2}, 2, 4, 15, {A1.MaxRows - 1}, 4");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertRows(A1.MaxRows - 1));

                Assert.Contains("comment note anchor", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.True(sheet.HasComment(1, 2));
            }
        }

        [Fact]
        public void Test_NamedStyleApplication_CancellationRollsBackPartialRange() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Style source");
            sheet.CellAt(1, 1).SetFillColor("C6EFCE");
            sheet.DefineNamedStyle("Input", 1, 1);
            sheet.CellValue(1, 26, "Probe");
            sheet.ApplyNamedStyle("Input", "Z1:Z1");
            uint appliedStyle = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "Z1").StyleIndex!.Value;
            sheet.CellValue(1, 2, "Keep");
            Cell observed = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            uint originalStyle = observed.StyleIndex?.Value ?? 0U;
            using var cancellation = new CancellationTokenSource();
            using var ready = new ManualResetEventSlim();
            var canceller = new Thread(() => {
                ready.Set();
                while ((observed.StyleIndex?.Value ?? 0U) != appliedStyle) Thread.Yield();
                cancellation.Cancel();
            }) { IsBackground = true };
            canceller.Start();
            ready.Wait();

            Assert.ThrowsAny<OperationCanceledException>(() => sheet.ApplyNamedStyle(
                "Input",
                "B1:XFD2",
                maximumCells: 40_000,
                cancellation.Token));
            Assert.True(canceller.Join(TimeSpan.FromSeconds(5)));

            Cell restored = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            Assert.Equal(originalStyle, restored.StyleIndex?.Value ?? 0U);
            Assert.DoesNotContain(sheet.WorksheetPart.Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "XFD1");
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernCharts_RejectNonFiniteValuesBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            foreach (double value in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity }) {
                var invalid = new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { value }) });
                Assert.Throws<ArgumentException>(() => sheet.AddModernChart(
                    invalid,
                    1,
                    1,
                    ExcelModernChartType.Funnel));
            }
            Assert.Empty(sheet.ModernCharts);
            Assert.Single(document.Sheets);

            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            ExcelChartDataRange originalRange = chart.DataRange!;
            Assert.Throws<ArgumentException>(() => chart.UpdateData(new ExcelChartData(
                new[] { "A" },
                new[] { new ExcelChartSeries("Value", new[] { double.NaN }) })));

            Assert.Equal(originalRange.StartRow, chart.DataRange!.StartRow);
            Assert.Equal(1d, document[originalRange.SheetName]
                .CellAt(originalRange.CategoryStartRow, originalRange.SeriesStartColumn).GetValue<double>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_QueryBackedTable_SchemaChangesSynchronizeNativeFields() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "SchemaQuery",
                WorksheetName = sheet.Name,
                TableName = "SchemaResults",
                ColumnNames = new[] { "A", "B" }
            });

            sheet.SetTableSchema(source.TableName, new[] { "Renamed", "B", "Added" }, "A1:C1");
            AssertQueryFieldSchema(sheet, new[] { "Renamed", "B", "Added" });

            sheet.ResizeTable(source.TableName, "A1:B1");
            AssertQueryFieldSchema(sheet, new[] { "Renamed", "B" });
            Assert.Empty(document.ValidateOpenXml());
        }

        private static void AssertQueryFieldSchema(ExcelSheet sheet, string[] expectedNames) {
            TableDefinitionPart tablePart = Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            TableColumn[] columns = tablePart.Table!.TableColumns!.Elements<TableColumn>().ToArray();
            QueryTableFields fields = Assert.Single(tablePart.QueryTableParts).QueryTable!
                .QueryTableRefresh!.QueryTableFields!;
            QueryTableField[] queryFields = fields.Elements<QueryTableField>().ToArray();
            Assert.Equal((uint)expectedNames.Length, fields.Count!.Value);
            Assert.Equal(expectedNames, queryFields.Select(field => field.Name!.Value));
            Assert.Equal(queryFields.Select(field => field.Id!.Value),
                columns.Select(column => column.QueryTableFieldId!.Value));
            Assert.Equal(columns.Select(column => column.Id!.Value),
                queryFields.Select(field => field.TableColumnId!.Value));
        }
    }
}
