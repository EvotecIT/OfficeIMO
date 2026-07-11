using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentCurrentReviewLossReportTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void WordToOdtReportsFlattenedNestedListLevels() {
        using WordDocument source = WordDocument.Create();
        WordList list = source.AddListNumbered();
        list.AddItem("Parent");
        list.AddItem("Child", 1);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocument();
        using OdtDocument target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "list-levels" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordReportsHeaderAndFooterImagesAsSkipped() {
        using OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Logo").AddImage(TinyPng, "header.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocument();
        using WordDocument target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void OdpToPowerPointReportsFlattenedListsAndMixedRuns() {
        using OdpPresentation source = OdpPresentation.Create();
        OdpTextBox textBox = source.AddSlide("Text").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 4), null, "Content");
        OdpParagraph mixed = textBox.AddParagraph("Plain ");
        mixed.AddRun("Bold").Bold = true;
        textBox.AddList().AddItem("Bullet");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentation();
        using PowerPointPresentation target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "text-lists" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains("Plain Bold", target.Slides.Single().TextBoxes.Single().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OdtToWordToleratesMissingStylesAndReportsTableCellImages() {
        using OdtDocument template = OdtDocument.Create();
        template.AddParagraph("Minimal");
        template.AddTable(1, 1, "Media").Cell(0, 0).Paragraphs[0].AddImage(TinyPng, "cell.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using OdtDocument source = OdtDocument.Open(new MemoryStream(RemovePackageEntry(template.ToBytes(), "styles.xml")));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocument();
        using WordDocument target = conversion.Document;

        Assert.Contains(target.CreateInspectionSnapshot().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphSnapshot>(), paragraph => paragraph.Text == "Minimal");
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void ExcelToOdsPreservesTypedValuesOnHyperlinkedCellsAndFormulaSeparators() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        ExcelSheet sheet = source.AddWorkSheet("Data");
        source.AddWorkSheet("Other, Sheet").CellAt(1, 1).SetValue(1);
        sheet.SetHyperlink(1, 1, "https://example.com", "42");
        sheet.CellAt(1, 1).SetValue(42);
        sheet.CellAt(1, 2).SetFormula("IF(A1=42,\"x,y\",\"other\")");
        sheet.CellAt(1, 3).SetFormula("SUM('Other, Sheet'!A1,A1)");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocument();
        using OdsDocument target = conversion.Document;
        OdsSheet converted = target.GetSheet("Data")!;

        Assert.Equal(OdsCellValueKind.Number, converted.GetValue(0, 0).Kind);
        Assert.Equal(42D, converted.GetValue(0, 0).AsDouble());
        Assert.Equal("https://example.com", converted.RowRuns[0].CellRuns[0].HyperlinkHref);
        Assert.Equal("of:=IF([.A1]=42;\"x,y\";\"other\")", converted.GetFormula(0, 1));
        Assert.Equal("of:=SUM([$'Other, Sheet'.A1];[.A1])", converted.GetFormula(0, 2));
    }

    [Fact]
    public void OdsToExcelCreatesInternalLinksWithoutLosingTypedValues() {
        using OdsDocument source = OdsDocument.Create();
        source.AddSheet("Target").Cell(0, 0).SetString("Destination");
        OdsCell linked = source.AddSheet("Links").Cell(0, 0);
        linked.SetNumber(42D);
        linked.SetHyperlink("Go", "#$'Target'.A1");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocument();
        using ExcelDocument target = conversion.Document;
        ExcelWorksheetSnapshot links = target.CreateInspectionSnapshot().Worksheets.Single(sheet => sheet.Name == "Links");
        ExcelCellSnapshot cell = Assert.Single(links.Cells);

        Assert.Equal(42m, Convert.ToDecimal(cell.Value));
        Assert.NotNull(cell.Hyperlink);
        Assert.False(cell.Hyperlink!.IsExternal);
        Assert.Equal("'Target'!A1", cell.Hyperlink.Target);
    }

    private static byte[] RemovePackageEntry(byte[] packageBytes, string removedPath) {
        using var input = new MemoryStream(packageBytes, writable: false);
        using var output = new MemoryStream();
        using (var source = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry sourceEntry in source.Entries.Where(entry => entry.FullName != removedPath)) {
                ZipArchiveEntry targetEntry = target.CreateEntry(sourceEntry.FullName,
                    sourceEntry.FullName == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
                using Stream sourceStream = sourceEntry.Open();
                using Stream targetStream = targetEntry.Open();
                sourceStream.CopyTo(targetStream);
            }
        }
        return output.ToArray();
    }
}
