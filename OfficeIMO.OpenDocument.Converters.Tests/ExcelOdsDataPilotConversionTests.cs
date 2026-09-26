using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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
    public void OdsPivotCannotWriteBeyondItsDeclaredTargetRange() {
        OdsDocument source = CreateSimpleOdsPivot("SmallTarget", "Data.D1:Data.D1");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void DuplicateOdsAxisFieldIsExplicitLoss() {
        OdsDocument source = CreateSimpleOdsPivot("DuplicateAxis", "Data.D1:Data.G3");
        source.DataPilotTables.Single().AddField("Region", "column");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Theory]
    [InlineData("sales")]
    [InlineData(" Sales ")]
    public void StaleOdsFieldBindingIsExplicitLoss(string staleName) {
        OdsDocument source = CreateSimpleOdsPivot("StaleBinding", "Data.D1:Data.E3");
        XElement field = source.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "data-pilot-field").Last();
        field.SetAttributeValue(OdfNamespaces.Table + "source-field-name", staleName);
        source.MarkPartDirty("content.xml");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void RetainedTargetRangesShareThePivotScanBudget() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        foreach ((string name, string targetRange) in new[] {
            ("First", "Data.D1:Data.E350000"),
            ("Second", "Data.G1:Data.H350000")
        }) {
            OdsDataPilotTable pivot = source.AddDataPilotTable(name, "Data.A1:Data.B2", targetRange);
            pivot.AddField("Region", "row");
            pivot.AddField("Sales", "data", "sum");
        }

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Single(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void RejectedOversizedPivotDoesNotConsumeTheNextPivotBudget() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable rejected = source.AddDataPilotTable("Rejected",
            "Data.A1:Data.B450000", "Data.D1:Data.E100001");
        rejected.AddField("Region", "row");
        rejected.AddField("Sales", "data", "sum");
        OdsDataPilotTable accepted = source.AddDataPilotTable("Accepted",
            "Data.A1:Data.B2", "Data.G1:Data.H75000");
        accepted.AddField("Region", "row");
        accepted.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Equal("Accepted", Assert.Single(target.Sheets.Single().GetPivotTables()).Name);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void DuplicateExcelAxisFieldIndexIsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        byte[] bytes = source.ToBytes();
        using (var stream = new MemoryStream(bytes)) {
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, true)) {
                var definition = package.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
                definition.RowFields!.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Field { Index = 0 });
                definition.RowFields.Count = 2;
                definition.Save();
            }
            bytes = stream.ToArray();
        }
        using ExcelDocument imported = ExcelDocument.Load(new MemoryStream(bytes));

        OdfConversionResult<OdsDocument> conversion = imported.ToOpenDocumentResult();
        Assert.Empty(conversion.Value.DataPilotTables);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void MissingTrailingOdsHeaderIsExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable("MissingHeader",
            "Data.A1:Data.C2", "Data.E1:Data.F3");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void PivotNameThatExcelWouldTrimIsExplicitLoss() {
        OdsDocument source = CreateSimpleOdsPivot(" PivotName ", "Data.D1:Data.E3");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Empty(target.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    private static OdsDocument CreateSimpleOdsPivot(string name, string targetRange) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable(name, "Data.A1:Data.B2", targetRange);
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");
        return source;
    }

    [Fact]
    public void DuplicateHeaderAddedAfterPivotCreationRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(1, 3, "Extra");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.CellValue(2, 3, "Note");
        sheet.AddPivotTable("A1:C2", "E1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        sheet.CellValue(1, 3, "Region");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        Assert.Empty(conversion.Value.DataPilotTables);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void DuplicateExcelPivotNamesOnDifferentSheetsRemainExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        foreach (string name in new[] { "North", "South" }) {
            ExcelSheet sheet = source.AddWorksheet(name);
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, name);
            sheet.CellValue(2, 2, 10d);
            sheet.AddPivotTable("A1:B2", "D1", name: "SharedName",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        }

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        Assert.Single(conversion.Value.DataPilotTables);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void CaseAmbiguousOdsHeadersRemainExplicitPivotLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Sales");
        sheet.Cell(0, 1).SetString("sales");
        sheet.Cell(0, 2).SetString("Region");
        sheet.Cell(1, 0).SetNumber(10);
        sheet.Cell(1, 1).SetNumber(20);
        sheet.Cell(1, 2).SetString("North");
        OdsDataPilotTable pivot = source.AddDataPilotTable("Ambiguous",
            "Data.A1:Data.C2", "Data.E1:Data.F3");
        pivot.AddField("Region", "row");
        pivot.AddField("sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument result = conversion.Value;
        Assert.Empty(result.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void WhitespaceAmbiguousOdsHeadersRemainExplicitPivotLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Sales");
        sheet.Cell(0, 1).SetString(" Sales ");
        sheet.Cell(0, 2).SetString("Region");
        sheet.Cell(1, 0).SetNumber(10);
        sheet.Cell(1, 1).SetNumber(20);
        sheet.Cell(1, 2).SetString("North");
        OdsDataPilotTable pivot = source.AddDataPilotTable("Ambiguous",
            "Data.A1:Data.C2", "Data.E1:Data.F3");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument result = conversion.Value;
        Assert.Empty(result.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void PivotWithStaleCacheHeaderCasingRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        sheet.CellValue(1, 1, "region");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        Assert.Empty(conversion.Value.DataPilotTables);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.DoesNotContain(conversion.Report.ForFeature("source-spreadsheet-data-pilot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void PivotAndNamedRangeKeepOdfSpreadsheetChildOrder() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        source.SetNamedRange("SalesValues", "'Data'!$B$2:$B$2", save: false);

        OdsDocument converted = source.ToOpenDocumentResult().Value;
        XElement spreadsheet = converted.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Office + "spreadsheet").Single();
        XName[] children = spreadsheet.Elements().Select(child => child.Name).ToArray();
        Assert.True(Array.IndexOf(children, OdfNamespaces.Table + "named-expressions")
            < Array.IndexOf(children, OdfNamespaces.Table + "data-pilot-tables"));
        Assert.Single(converted.NamedRanges);
        Assert.Single(converted.DataPilotTables);
        Assert.True(converted.Validate().IsValid);
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

    [Fact]
    public void PivotWhoseGeneratedLocationExceedsWorksheetRemainsExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable("EdgePivot",
            "Data.A1:Data.B2", "Data.XFD1048576:Data.XFD1048576");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument result = conversion.Value;
        Assert.Empty(result.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void EmptyRepeatedTailOutsidePivotDoesNotDiscardItsConvertedRange() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable("SalesPivot", "Data.A1:Data.B2", "Data.D1:Data.E3");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");
        XElement table = source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table").Single();
        table.Add(new XElement(OdfNamespaces.Table + "table-row",
            new XAttribute(OdfNamespaces.Table + "number-rows-repeated", 1_100_000),
            new XElement(OdfNamespaces.Table + "table-cell")));
        source.MarkPartDirty("content.xml");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument result = conversion.Value;
        Assert.Single(result.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("expansion-limits"),
            mapping => mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void PivotWithClippedSourceCellRemainsExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = source.AddDataPilotTable("SalesPivot", "Data.A1:Data.B2", "Data.D1:Data.E3");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { MaximumExpandedCells = 3 });
        using ExcelDocument result = conversion.Value;
        Assert.Empty(result.Sheets.Single().GetPivotTables());
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void ExcelValuesAxisSentinelRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        byte[] bytes = source.ToBytes();
        using (var stream = new MemoryStream(bytes)) {
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, true)) {
                var definition = package.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
                definition.RowFields!.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Field { Index = -2 });
                definition.RowFields.Count = 2;
                definition.Save();
            }
            bytes = stream.ToArray();
        }
        using ExcelDocument imported = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Contains("Field-1", imported.Sheets.Single().GetPivotTables().Single().RowFields);
        Assert.True(imported.Sheets.Single().GetPivotTables().Single().HasValuesAxisField);

        OdfConversionResult<OdsDocument> conversion = imported.ToOpenDocumentResult();
        Assert.Empty(conversion.Value.DataPilotTables);
        Assert.Contains(conversion.Report.ForFeature("pivot-tables"),
            mapping => mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void RealHeaderNamedLikeValuesSentinelStillConverts() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Field-1");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 10d);
        sheet.AddPivotTable("A1:B2", "D1", name: "SalesPivot",
            rowFields: new[] { "Field-1" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
        Assert.False(sheet.GetPivotTables().Single().HasValuesAxisField);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDataPilotTable pivot = Assert.Single(conversion.Value.DataPilotTables);
        Assert.Equal("Field-1", pivot.Fields[0].SourceFieldName);
    }
}
