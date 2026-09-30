using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_GoogleSheetsPivot_ReportsUnsupportedMiddleValuesPlacement(bool valuesOnRows) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Product"); sheet.CellValue(1, 3, "Amount");
            sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, "A"); sheet.CellValue(2, 3, 10d);
            var fields = new[] { "Region", "Product" };
            sheet.AddPivotTable("A1:C2", "E1", "Pivot",
                rowFields: valuesOnRows ? fields : Array.Empty<string>(),
                columnFields: valuesOnRows ? Array.Empty<string>() : fields,
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue"),
                    new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Count, "Entries") }, dataOnRows: valuesOnRows);
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            var axis = valuesOnRows ? (DocumentFormat.OpenXml.OpenXmlCompositeElement)definition.RowFields! : definition.ColumnFields!;
            axis.RemoveAllChildren();
            axis.Append(new Field { Index = 0 }, new Field { Index = -2 }, new Field { Index = 1 });
            var metadata = Assert.Single(sheet.GetPivotTables());
            Assert.Equal(1, metadata.ValuesAxisPosition);
            Assert.Equal(fields, valuesOnRows ? metadata.RowSourceFields : metadata.ColumnSourceFields);
            var batch = document.BuildGoogleSheetsBatch();
            Assert.Empty(batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>());
            Assert.Single(batch.Report.Notices, n => n.Code == "SHEETS.PIVOT_TABLE.UNSUPPORTED");
        }

        [Theory]
        [InlineData(false, "HORIZONTAL", ExcelPivotTableAxis.AxisColumn)]
        [InlineData(true, "VERTICAL", ExcelPivotTableAxis.AxisRow)]
        public void Test_GoogleSheetsPivot_PreservesMeasureOrientationAndRealValuesNamedField(bool valuesOnRows, string layout, ExcelPivotTableAxis axis) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Values"); sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, 10d);
            sheet.AddPivotTable("A1:B2", "D1", "Pivot", rowFields: new[] { "Values" },
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue"),
                    new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Count, "Entries") }, dataOnRows: valuesOnRows);
            var metadata = Assert.Single(sheet.GetPivotTables());
            Assert.Equal(new[] { "Values" }, metadata.RowSourceFields);
            Assert.Empty(metadata.ColumnSourceFields);
            Assert.Equal(axis, metadata.ValuesAxis);
            Assert.Equal(valuesOnRows ? 1 : 0, metadata.ValuesAxisPosition);
            var batch = document.BuildGoogleSheetsBatch();
            var request = Assert.Single(batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>());
            Assert.Equal(0, Assert.Single(request.Rows).SourceColumnOffset);
            Assert.Empty(request.Columns);
            Assert.Equal(layout, request.ValueLayout);
            var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
            var pivot = payload.Requests.Where(r => r.UpdateCells != null).SelectMany(r => r.UpdateCells!.Rows)
                .SelectMany(r => r.Values).Select(c => c.PivotTable).Single(p => p != null)!;
            Assert.Equal(layout, pivot.ValueLayout);
            using var json = JsonDocument.Parse(JsonSerializer.Serialize(pivot));
            Assert.Equal(layout, json.RootElement.GetProperty("valueLayout").GetString());
        }

        [Fact]
        public void Test_GoogleSheetsPivot_ReportsUnsupportedOuterValuesPlacement() {
            using var document = ExcelDocument.Load(PivotLayoutOraclePath);
            var batch = document.BuildGoogleSheetsBatch();
            var emitted = batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>().ToArray();
            Assert.Equal(8, emitted.Length);
            Assert.DoesNotContain(emitted, p => p.SheetName.EndsWith("Outer", StringComparison.Ordinal));
            Assert.Equal(4, batch.Report.Notices.Count(n => n.Code == "SHEETS.PIVOT_TABLE.UNSUPPORTED"));
        }
    }
}
