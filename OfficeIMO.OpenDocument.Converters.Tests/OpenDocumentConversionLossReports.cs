using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellFormula = DocumentFormat.OpenXml.Spreadsheet.CellFormula;
using OpenXmlCellFormulaValues = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues;
using OpenXmlSpreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentConversionLossReportTests {
    [Fact]
    public void ExcelToOdsReportsOfficeImoPivotBindingMetadataLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet data = source.AddWorksheet("Data");
        data.CellValue(1, 1, "Region");
        data.CellValue(1, 2, "OrderDate");
        data.CellValue(1, 3, "Sales");
        data.CellValue(2, 1, "East");
        data.CellValue(2, 2, new DateTime(2026, 1, 2));
        data.CellValue(2, 3, 10d);
        data.AddPivotTable(
            sourceRange: "A1:C2",
            destinationCell: "E2",
            name: "SalesPivot",
            rowFields: new[] { "Region" });
        data.AddChart(
            new ExcelChartData(
                new[] { "East" },
                new[] { new ExcelChartSeries("Sales", new[] { 10d }) }),
            row: 8,
            column: 1,
            widthPixels: 500,
            heightPixels: 300,
            type: ExcelChartType.ColumnClustered,
            title: "Sales");
        source.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions {
            Name = "RegionSlicer",
            SourceName = "Region"
        });
        source.AddWorkbookTimelineCache(new ExcelTimelineCacheOptions {
            Name = "OrderTimeline",
            SourceName = "OrderDate"
        });

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        ExcelWorkbookSnapshot snapshot = source.CreateInspectionSnapshot();
        Assert.Equal(1, snapshot.ChartPartCount);
        Assert.Equal(1, snapshot.PivotTablePartCount);
        Assert.True(snapshot.HasCharts);
        Assert.True(snapshot.HasPivotTables);

        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "slicer-binding-metadata"
            && mapping.Status == OdfConversionMappingStatus.Unsupported
            && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "timeline-binding-metadata"
            && mapping.Status == OdfConversionMappingStatus.Unsupported
            && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "charts"
            && mapping.Status == OdfConversionMappingStatus.Unsupported
            && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "pivot-tables"
            && mapping.Status == OdfConversionMappingStatus.Unsupported
            && mapping.Count == 1);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx, "XLSX")]
    [InlineData(ExcelFileFormat.Xls, "XLS")]
    [InlineData(ExcelFileFormat.Xlsb, "XLSB")]
    public void ExcelToOdsReportUsesActualSourceFormat(ExcelFileFormat format, string expectedLabel) {
        byte[] bytes;
        using (ExcelDocument created = ExcelDocument.Create()) {
            created.AddWorksheet("Data").CellValue(1, 1, 1d);
            bytes = created.ToBytes(format);
        }

        using ExcelDocument source = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Equal(format, source.SourceFormat);
        Assert.Equal(expectedLabel, conversion.Report.SourceFormat);
    }

    [Fact]
    public void ExcelFormulaConversionDoesNotRewriteQuotedCellLikeText() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue("B2");
        sheet.CellAt(1, 2).SetFormula("IF(A1=\"B2\",1,0)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDocument target = conversion.Value;

        Assert.Equal("of:=IF([.A1]=\"B2\";1;0)", target.GetSheet("Data")!.Cell(0, 1).Formula);

        target.GetSheet("Data")!.Cell(0, 2).Formula = "of:=IF([.A1]=\"[.B2]\",1,0)";
        OdfConversionResult<ExcelDocument> reverse = target.ToExcelDocumentResult();
        using ExcelDocument roundTrip = reverse.Value;
        ExcelCellSnapshot reverseFormula = roundTrip.CreateInspectionSnapshot().Worksheets.Single().Cells
            .Single(cell => cell.Row == 1 && cell.Column == 3);
        Assert.Equal("IF(A1=\"[.B2]\",1,0)", reverseFormula.Formula);
    }

    [Fact]
    public void ExcelFormulaConversionMapsSheetQualifiedReferences() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet data = source.AddWorksheet("Data");
        source.AddWorksheet("Other").CellAt(1, 1).SetValue(1);
        source.AddWorksheet("Other Sheet").CellAt(2, 2).SetValue(2);
        data.CellAt(1, 1).SetFormula("SUM(Other!A1,'Other Sheet'!B2:C3)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDocument target = conversion.Value;

        Assert.Equal("of:=SUM([$'Other'.A1];[$'Other Sheet'.B2:.C3])",
            target.GetSheet("Data")!.Cell(0, 0).Formula);
        OdfConversionResult<ExcelDocument> reverse = target.ToExcelDocumentResult();
        using ExcelDocument roundTrip = reverse.Value;
        ExcelCellSnapshot formula = roundTrip.CreateInspectionSnapshot().Worksheets.Single(sheet => sheet.Name == "Data").Cells.Single();
        Assert.Equal("SUM('Other'!A1,'Other Sheet'!B2:C3)", formula.Formula);
    }

    [Fact]
    public void ExcelToOdsExpandsSharedFormulaFollowers() {
        string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellFormula(1, 2, "A1+1");
                sheet.CellFormula(2, 2, "A2+1");
                document.Save();
            }

            using (OpenXmlSpreadsheetDocument spreadsheet = OpenXmlSpreadsheetDocument.Open(filePath, true)) {
                OpenXmlWorksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet
                    ?? throw new InvalidDataException("Worksheet is missing.");
                OpenXmlCell master = Assert.Single(worksheet.Descendants<OpenXmlCell>(), cell => cell.CellReference?.Value == "B1")!;
                OpenXmlCell follower = Assert.Single(worksheet.Descendants<OpenXmlCell>(), cell => cell.CellReference?.Value == "B2")!;
                master.CellFormula = new OpenXmlCellFormula("A1+1") {
                    FormulaType = OpenXmlCellFormulaValues.Shared,
                    SharedIndex = 0U,
                    Reference = "B1:B2"
                };
                follower.CellFormula = new OpenXmlCellFormula {
                    FormulaType = OpenXmlCellFormulaValues.Shared,
                    SharedIndex = 0U
                };
                worksheet.Save();
            }

            using ExcelDocument source = ExcelDocument.Load(filePath);
            OdsDocument converted = source.ToOpenDocument();
            Assert.Equal("of:=[.A2]+1", converted.GetSheet("Data")!.Cell(1, 1).Formula);
        } finally {
            if (File.Exists(filePath)) {
                File.Delete(filePath);
            }
        }
    }

    [Fact]
    public void OpenFormulaConversionMapsArgumentSeparatorsOutsideStrings() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).SetNumber(1);
        sheet.Cell(0, 1).Formula = "of:=IF([.A1]>0;\"yes;still\";\"no\")";

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        ExcelCellSnapshot formula = target.CreateInspectionSnapshot().Worksheets.Single().Cells
            .Single(cell => cell.Row == 1 && cell.Column == 2);
        Assert.Equal("IF(A1>0,\"yes;still\",\"no\")", formula.Formula);
    }

    [Fact]
    public void DuplicateSheetLocalExcelNamesAreDisambiguatedWithoutAborting() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet first = source.AddWorksheet("First Sheet");
        ExcelSheet second = source.AddWorksheet("Second Sheet");
        first.CellAt(1, 1).SetValue(1);
        second.CellAt(1, 1).SetValue(2);
        first.SetNamedRange("LocalValue", "A1", save: false);
        second.SetNamedRange("LocalValue", "A1", save: false);

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDocument target = conversion.Value;

        Assert.Equal(2, target.NamedRanges.Count);
        Assert.Equal(2, target.NamedRanges.Select(named => named.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(target.NamedRanges, named => named.CellRangeAddress.Contains("First Sheet"));
        Assert.Contains(target.NamedRanges, named => named.CellRangeAddress.Contains("Second Sheet"));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "sheet-local-named-ranges" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void WordAutomaticColorsAndUnsupportedImagesDoNotAbortConversion() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Automatic color").ColorHex = "auto";
        using var tiff = new MemoryStream(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00 });
        source.AddParagraph().AddImage(tiff, "unsupported.tiff", 10, 10);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtDocument target = conversion.Value;

        Assert.Equal("Automatic color", target.Paragraphs.First().Text);
        Assert.Null(target.Paragraphs.First().Spans.Single().Color);
        Assert.Empty(target.Paragraphs.SelectMany(paragraph => paragraph.Images));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordPreservesRelativeLinksAndSkipsUnsupportedImages() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddHyperlink("Relative", "docs/page.html");
        byte[] webp = { 0x52, 0x49, 0x46, 0x46, 0x04, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50 };
        source.AddParagraph("Image").AddImage(webp, "pixel.webp", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;

        WordRunSnapshot link = target.CreateInspectionSnapshot().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphSnapshot>().Single(paragraph => paragraph.Text == "Relative").Runs.Single();
        Assert.Equal("docs/page.html", link.HyperlinkUri);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void UnavailableOdfImagesAreReportedWithoutAbortingOfficeExport() {
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace xlink = "http://www.w3.org/1999/xlink";
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

        OdtDocument textSource = OdtDocument.Create();
        textSource.AddParagraph().AddImage(png, "missing.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        OdtDocument brokenText = OpenBrokenFlat(textSource.ToFlatXml(), draw, xlink, office,
            stream => OdtDocument.LoadFlatXml(stream));
        OdfConversionResult<WordDocument> wordConversion = brokenText.ToWordDocumentResult();
        using WordDocument word = wordConversion.Value;

        Assert.Contains(wordConversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);

        OdpPresentation presentationSource = OdpPresentation.Create();
        presentationSource.AddSlide("Broken").AddImage(png, "missing.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        OdpPresentation brokenPresentation = OpenBrokenFlat(presentationSource.ToFlatXml(), draw, xlink, office,
            stream => OdpPresentation.LoadFlatXml(stream));
        OdfConversionResult<PowerPointPresentation> powerPointConversion = brokenPresentation.ToPowerPointPresentationResult();
        using PowerPointPresentation powerPoint = powerPointConversion.Value;

        Assert.Contains(powerPointConversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void UnsupportedPowerPointImageFormatsAreReportedWithoutAbortingConversion() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointSlide slide = source.AddSlide();
        using var tiff = new MemoryStream(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x00 });
        slide.AddPicture(tiff, ImagePartType.Tiff);

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpPresentation target = conversion.Value;

        Assert.Empty(Assert.Single(target.Slides).Shapes.OfType<OdpImage>());
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdtTrackedChangesAreReportedWhenConvertingToWord() {
        OdtDocument source = OdtDocument.Create();
        source.AddTrackedParagraphInsertion("Inserted", "Author");

        OdfConversionResult<OfficeIMO.Word.WordDocument> conversion = source.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument target = conversion.Value;

        Assert.True(conversion.Report.HasLoss);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-tracked-changes" && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdpAnimationsAreReportedWhenConvertingToPowerPoint() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Animated");
        OdpRectangle shape = slide.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        slide.AddFadeInAnimation(shape, TimeSpan.FromSeconds(1));

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;

        Assert.True(conversion.Report.HasLoss);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-presentation-animations" && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdsExpansionLimitStillCreatesEveryWorksheetAndReportsTruncation() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet hidden = source.AddSheet("Hidden");
        hidden.Hidden = true;
        hidden.Cell(0, 0).SetNumber(1);
        hidden.Cell(0, 1).SetNumber(2);
        source.AddSheet("Visible").Cell(0, 0).SetNumber(3);

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult(new ExcelOpenDocumentConversionOptions {
            MaximumExpandedCells = 1
        });
        using ExcelDocument target = conversion.Value;

        Assert.Equal(2, target.CreateInspectionSnapshot().Worksheets.Count);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "worksheets" && mapping.Count == 2);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "expansion-limits" && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void OdsToExcelKeepsOneWorksheetVisibleWhenEverySourceSheetIsHidden() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("First").Hidden = true;
        source.AddSheet("Second").Hidden = true;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelWorkbookSnapshot snapshot = target.CreateInspectionSnapshot();

        Assert.Empty(target.ValidateDocument());
        Assert.False(snapshot.Worksheets[0].Hidden);
        Assert.True(snapshot.Worksheets[0].IsActive);
        Assert.True(snapshot.Worksheets[1].Hidden);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "worksheet-visibility" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void ExcelToOdsReportsConfiguredCellAndStyleOmissions() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue("One").SetBold();
        sheet.CellAt(1, 2).SetValue("Two").SetBold();

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult(new ExcelOpenDocumentConversionOptions {
            IncludeBasicStyles = false,
            MaximumExpandedCells = 1
        });
        OdsDocument target = conversion.Value;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "cell-styles" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "expansion-limits" && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    private static T OpenBrokenFlat<T>(XDocument flat, XNamespace draw, XNamespace xlink, XNamespace office,
        Func<Stream, T> load) {
        XElement image = flat.Descendants(draw + "image").Single();
        image.SetAttributeValue(xlink + "href", "https://example.test/missing.png");
        image.Elements(office + "binary-data").Remove();
        var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        try {
            return load(stream);
        } finally {
            stream.Dispose();
        }
    }
}
