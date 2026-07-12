using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentConversionContracts {
    [Fact]
    public void WordAndOdtRoundTripSemanticBlocksAndReportFidelity() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Native OpenDocument conversion").Style = WordParagraphStyles.Heading1;
        source.AddParagraph("Paragraph with structured content.").Bold = true;
        source.AddListNumbered().AddItem("First item");
        WordTable sourceTable = source.AddTable(2, 2);
        sourceTable.Rows[0].Cells[0].Paragraphs[0].Text = "A";
        sourceTable.Rows[1].Cells[1].Paragraphs[0].Text = "B";

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocument();
        using OdtDocument odt = toOdt.Document;
        Assert.True(odt.Validate().IsValid);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "headings");
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "tables");

        using var package = new MemoryStream(odt.ToBytes());
        using OdtDocument reopened = OdtDocument.Open(package);
        Assert.Contains(reopened.ContentBlocks, block => block.Paragraph?.Text == "Native OpenDocument conversion");
        Assert.Contains(reopened.ContentBlocks, block => block.Table != null);

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocument();
        using WordDocument roundTrip = toWord.Document;
        Assert.Empty(roundTrip.ValidateDocument());
        WordDocumentSnapshot snapshot = roundTrip.CreateInspectionSnapshot();
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>(),
            paragraph => paragraph.Text.Contains("First item", StringComparison.Ordinal));
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements), block => block is WordTableSnapshot);
    }

    [Fact]
    public void ExcelAndOdsRoundTripTypedCellsFormulaMergeAndSparseLimits() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorkSheet("Data");
        sheet.CellAt(1, 1).SetValue("Amount").SetBold();
        sheet.CellAt(2, 1).SetValue(12.5m);
        sheet.CellAt(2, 2).SetFormula("SUM(A2:A2)");
        sheet.SetHyperlink(3, 1, "https://example.com", "Example");
        sheet.MergeRange("A4:B4");
        sheet.CellAt(4, 1).SetValue("Merged");
        source.SetNamedRange("Amounts", "'Data'!$A$2:$A$2", save: false);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocument();
        using OdsDocument ods = toOds.Document;
        Assert.True(ods.Validate().IsValid);
        Assert.Equal(12.5m, ods.GetSheet("Data")!.Cell(1, 0).Value.AsDecimal());
        Assert.StartsWith("of:=", ods.GetSheet("Data")!.Cell(1, 1).Formula);
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "formulas" && mapping.Status == OdfConversionMappingStatus.Approximated);

        using var package = new MemoryStream(ods.ToBytes());
        using OdsDocument reopened = OdsDocument.Open(package);
        OdfConversionResult<ExcelDocument> toExcel = reopened.ToExcelDocument(new ExcelOpenDocumentConversionOptions {
            MaximumExpandedCells = 1000
        });
        using ExcelDocument roundTrip = toExcel.Document;
        Assert.Empty(roundTrip.ValidateDocument());
        ExcelWorksheetSnapshot snapshot = Assert.Single(roundTrip.CreateInspectionSnapshot().Worksheets);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 1 && Convert.ToDecimal(cell.Value) == 12.5m);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 2 && cell.Formula != null);
        Assert.Contains(snapshot.MergedRanges, merge => merge.A1Range == "A4:B4");
    }

    [Fact]
    public void PowerPointAndOdpRoundTripSlidesShapesTablesNotesAndTransitions() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointStreamCreateOptions { AutoSave = false });
        PowerPointSlide slide = source.AddSlide();
        PowerPointTextBox title = slide.AddTextBoxPoints("OpenDocument deck", 24, 20, 240, 40);
        title.Paragraphs[0].Runs[0].Bold = true;
        slide.AddRectanglePoints(30, 80, 100, 50, "Panel").FillColor = "DDEEFF";
        PowerPointTable table = slide.AddTablePoints(2, 2, 150, 80, 160, 80);
        table.GetCell(0, 0).Text = "A";
        table.GetCell(1, 1).Text = "B";
        table.MergeCells(0, 0, 0, 1);
        slide.Notes.Text = "Speaker note";
        slide.Transition = SlideTransition.Fade;

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocument();
        using OdpPresentation odp = toOdp.Document;
        Assert.True(odp.Validate().IsValid);
        OdpSlide odpSlide = Assert.Single(odp.Slides);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTextBox);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTable);
        Assert.Equal("Speaker note", Assert.Single(odpSlide.SpeakerNotes!.Paragraphs).Text);
        Assert.Contains(toOdp.Report.Mappings, mapping => mapping.Feature == "slide-transitions");

        using var package = new MemoryStream(odp.ToBytes());
        using OdpPresentation reopened = OdpPresentation.Open(package);
        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopened.ToPowerPointPresentation();
        using PowerPointPresentation roundTrip = toPowerPoint.Document;
        Assert.Empty(roundTrip.ValidateDocument());
        PowerPointSlide roundTripSlide = Assert.Single(roundTrip.Slides);
        Assert.Contains(roundTripSlide.TextBoxes, box => box.Text.Contains("OpenDocument deck", StringComparison.Ordinal));
        Assert.Single(roundTripSlide.Tables);
        Assert.Equal("Speaker note", roundTripSlide.GetSpeakerNotesText());
        Assert.Equal(SlideTransition.Fade, roundTripSlide.Transition);
    }
}
