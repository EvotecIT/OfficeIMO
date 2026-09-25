using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class ExcelOdsDataPilotConversionTests {
    [Fact]
    public void BasicExcelPivotBecomesNativeDataPilotAndReopens() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        OdsDataPilotTable pivot = Assert.Single(reopened.DataPilotTables);
        Assert.Equal("SalesPivot", pivot.Name);
        Assert.Equal("row", pivot.Fields[0].Orientation);
        Assert.Equal("sum", pivot.Fields[1].Function);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void ExcelProducedPivotWithOrdinaryVisibleItemsBecomesDataPilot() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-pivot.xlsx");
        using ExcelDocument source = ExcelDocument.Load(path);
        ExcelPivotTableInfo original = Assert.Single(source.Sheets.Single().GetPivotTables());
        Assert.Contains(original.Fields, field => field.VisibleItems.Count > 0 && field.HiddenItems.Count == 0);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDataPilotTable mapped = Assert.Single(conversion.Value.DataPilotTables);
        Assert.Equal(original.Name, mapped.Name);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void ExcelProducedDataPilotBecomesWorksheetPivotWithExplicitApproximation() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-pivot.ods");
        OdsDocument source = OdsDocument.Load(path);

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelPivotTableInfo pivot = Assert.Single(target.Sheets.Single().GetPivotTables());
        Assert.Equal("SalesPivot", pivot.Name);
        Assert.Equal("A1:C5", pivot.SourceRange);
        Assert.Contains("Region", pivot.RowFields);
        Assert.Contains("Month", pivot.ColumnFields);
        Assert.Equal("Sales", Assert.Single(pivot.DataFields).FieldName);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(conversion.Report.ForFeature("source-spreadsheet-data-pilot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void DataPilotFollowsSheetWhenExcelSanitizesItsName() {
        const string sourceName = "WorksheetNameLongerThanThirtyOneCharacters";
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet(sourceName);
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable authored = source.AddDataPilotTable("RenamedPivot",
            sourceName + ".A1:" + sourceName + ".B2",
            sourceName + ".D1:" + sourceName + ".E3");
        authored.AddField("Region", "row");
        authored.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelSheet mappedSheet = Assert.Single(target.Sheets);
        Assert.NotEqual(sourceName, mappedSheet.Name);
        Assert.Equal("RenamedPivot", Assert.Single(mappedSheet.GetPivotTables()).Name);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void UnsupportedPageFieldRemainsExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable("Filtered", "Data.A1:Data.B2", "Data.D1:Data.E3");
        pivot.AddField("Region", "page");
        pivot.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }
}
