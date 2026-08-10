using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetWorkbookSemanticsConversionTests {
    [Fact]
    public void ExcelBackslashDefinedNameIsPreservedExactlyInOds() {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("Data");
        source.SetNamedRange("\\Rate", "'Data'!A1", save: false,
            validationMode: ExcelDefinedNameValidationMode.Strict);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        OdsNamedRange namedRange = Assert.Single(conversion.Value.NamedRanges);
        Assert.Equal("\\Rate", namedRange.Name);
        OdsDocument persisted = OdsDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        Assert.True(persisted.Validate().IsValid);
        Assert.Equal("\\Rate", Assert.Single(persisted.NamedRanges).Name);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping =>
            mapping.Feature == "sheet-local-named-ranges"
            && mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void ExcelNamedRangeHyperlinksFollowCollisionSafeOdsNames() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet north = source.AddWorksheet("North");
        ExcelSheet south = source.AddWorksheet("South");
        north.SetNamedRange("LocalData", "A1", save: false);
        south.SetNamedRange("LocalData", "B1", save: false);
        north.SetInternalLink(1, 2, "LocalData", "North data");
        south.SetInternalLink(1, 3, "LocalData", "South data");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Equal("#LocalData", FindHyperlink(conversion.Value.GetSheet("North")!, 0, 1));
        Assert.Equal("#LocalData__South", FindHyperlink(conversion.Value.GetSheet("South")!, 0, 2));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "sheet-local-named-ranges"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlinks"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ExcelHyperlinkTooltipIsInspectedAndReportedAsUnsupportedByOds() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.SetHyperlink(1, 1, "https://example.test/docs", "Docs", style: true, tooltip: "Open docs");

        ExcelCellSnapshot authored = Assert.Single(source.CreateInspectionSnapshot().Worksheets.Single().Cells);
        Assert.Equal("Open docs", authored.Hyperlink!.Tooltip);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Equal("https://example.test/docs", FindHyperlink(conversion.Value.Sheets.Single(), 0, 0));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "hyperlink-tooltips"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    [Fact]
    public void ValidationDropdownVisibilityRoundTripsThroughTypedOdfDisplayList() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.ValidationList("A1", new[] { "One", "Two" });
        sheet.SetDataValidationMessages("A1", new ExcelDataValidationMessageOptions {
            SuppressDropDown = true,
            PreserveShowMessageFlags = true
        });

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        Assert.Equal(OdsValidationDisplayList.None, Assert.Single(toOds.Value.Validations).DisplayList);

        OdsDocument persisted = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(persisted.Validate().IsValid);
        Assert.Equal(OdsValidationDisplayList.None, Assert.Single(persisted.Validations).DisplayList);
        using ExcelDocument roundTrip = persisted.ToExcelDocumentResult().Value;
        Assert.True(Assert.Single(roundTrip.Sheets.Single().GetDataValidations()).SuppressDropDown);
    }

    [Fact]
    public void SortedOdfValidationListIsReportedWhenExcelCannotPreserveItsDisplayOrder() {
        OdsDocument source = OdsDocument.Create();
        OdsValidation validation = source.AddValidation(
            "Names", OdsValidationConditionSyntax.CreateList(new[] { "B", "A" }));
        validation.DisplayList = OdsValidationDisplayList.SortAscending;
        source.AddSheet("Data").Cell(0, 0).ValidationName = validation.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        Assert.False(Assert.Single(target.Sheets.Single().GetDataValidations()).SuppressDropDown);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "validation-display-lists"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void SpreadsheetMetadataMapsCreatorAndSubjectInBothDirections() {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("Data");
        source.BuiltinDocumentProperties.Title = "Quarterly report";
        source.BuiltinDocumentProperties.Creator = "Alice";
        source.BuiltinDocumentProperties.Subject = "Forecast";

        OdsDocument ods = OdsDocument.Load(new MemoryStream(source.ToOpenDocument().ToBytes()));
        Assert.Equal("Quarterly report", ods.Metadata.Title);
        Assert.Equal("Alice", ods.Metadata.Creator);
        Assert.Equal("Forecast", ods.Metadata.Subject);

        using ExcelDocument roundTrip = ods.ToExcelDocument();
        Assert.Equal("Quarterly report", roundTrip.BuiltinDocumentProperties.Title);
        Assert.Equal("Alice", roundTrip.BuiltinDocumentProperties.Creator);
        Assert.Equal("Forecast", roundTrip.BuiltinDocumentProperties.Subject);
    }

    [Fact]
    public void NonFirstActiveExcelWorksheetIsReportedUnsupportedByOds() {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("First");
        ExcelSheet second = source.AddWorksheet("Second");
        source.SetActiveWorksheet(second);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "worksheet-views"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => conversion.Report.RequireNoSkippedOrUnsupported());
    }

    private static string? FindHyperlink(OdsSheet sheet, long row, long column) => sheet.RowRuns
        .Where(rowRun => row >= rowRun.StartRow && row < rowRun.StartRow + rowRun.RepeatCount)
        .SelectMany(rowRun => rowRun.CellRuns)
        .Single(cellRun => column >= cellRun.StartColumn && column < cellRun.StartColumn + cellRun.RepeatCount)
        .HyperlinkHref;
}
