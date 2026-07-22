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

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocumentResult();
        OdtDocument odt = toOdt.Value;
        Assert.True(odt.Validate().IsValid);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "headings");
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "tables");

        using var package = new MemoryStream(odt.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        Assert.Contains(reopened.ContentBlocks, block => block.Paragraph?.Text == "Native OpenDocument conversion");
        Assert.Contains(reopened.ContentBlocks, block => block.Table != null);

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocumentResult();
        using WordDocument roundTrip = toWord.Value;
        roundTrip.AddParagraph("Detached conversion remains editable");
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        WordDocumentSnapshot snapshot = roundTrip.CreateInspectionSnapshot();
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>(),
            paragraph => paragraph.Text.Contains("First item", StringComparison.Ordinal));
        Assert.Contains(snapshot.Sections.SelectMany(section => section.Elements), block => block is WordTableSnapshot);
    }

    [Fact]
    public void ExcelAndOdsRoundTripTypedCellsFormulaMergeAndSparseLimits() {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue("Amount").SetBold();
        sheet.CellAt(2, 1).SetValue(12.5m);
        sheet.CellAt(2, 2).SetFormula("SUM(A2:A2)");
        sheet.SetHyperlink(3, 1, "https://example.com", "Example");
        sheet.MergeRange("A4:B4");
        sheet.CellAt(4, 1).SetValue("Merged");
        source.SetNamedRange("Amounts", "'Data'!$A$2:$A$2", save: false);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        OdsDocument ods = toOds.Value;
        Assert.True(ods.Validate().IsValid);
        Assert.Equal(12.5m, ods.GetSheet("Data")!.Cell(1, 0).Value.AsDecimal());
        Assert.StartsWith("of:=", ods.GetSheet("Data")!.Cell(1, 1).Formula);
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "formulas" && mapping.Status == OdfConversionMappingStatus.Approximated);

        using var package = new MemoryStream(ods.ToBytes());
        OdsDocument reopened = OdsDocument.Load(package);
        OdfConversionResult<ExcelDocument> toExcel = reopened.ToExcelDocumentResult(new ExcelOpenDocumentConversionOptions {
            MaximumExpandedCells = 1000
        });
        using ExcelDocument roundTrip = toExcel.Value;
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        ExcelWorksheetSnapshot snapshot = Assert.Single(roundTrip.CreateInspectionSnapshot().Worksheets);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 1 && Convert.ToDecimal(cell.Value) == 12.5m);
        Assert.Contains(snapshot.Cells, cell => cell.Row == 2 && cell.Column == 2 && cell.Formula != null);
        Assert.Contains(snapshot.MergedRanges, merge => merge.A1Range == "A4:B4");
    }

    [Fact]
    public void PowerPointAndOdpRoundTripSlidesShapesTablesNotesAndTransitions() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
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

        OdfConversionResult<OdpPresentation> toOdp = source.ToOpenDocumentResult();
        OdpPresentation odp = toOdp.Value;
        Assert.True(odp.Validate().IsValid);
        OdpSlide odpSlide = Assert.Single(odp.Slides);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTextBox);
        Assert.Contains(odpSlide.Shapes, shape => shape is OdpTable);
        Assert.Equal("Speaker note", Assert.Single(odpSlide.SpeakerNotes!.Paragraphs).Text);
        Assert.Contains(toOdp.Report.Mappings, mapping => mapping.Feature == "slide-transitions");

        using var package = new MemoryStream(odp.ToBytes());
        OdpPresentation reopened = OdpPresentation.Load(package);
        OdfConversionResult<PowerPointPresentation> toPowerPoint = reopened.ToPowerPointPresentationResult();
        using PowerPointPresentation roundTrip = toPowerPoint.Value;
        Assert.Throws<InvalidOperationException>(() => roundTrip.Save());
        Assert.Empty(roundTrip.ValidateDocument());
        PowerPointSlide roundTripSlide = Assert.Single(roundTrip.Slides);
        Assert.Contains(roundTripSlide.TextBoxes, box => box.Text.Contains("OpenDocument deck", StringComparison.Ordinal));
        Assert.Single(roundTripSlide.Tables);
        Assert.Equal("Speaker note", roundTripSlide.GetSpeakerNotesText());
        Assert.Equal(SlideTransition.Fade, roundTripSlide.Transition);
    }

    [Fact]
    public void OdpToPowerPointRejectsTablesBeyondConfiguredBounds() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Bounded");
        slide.AddTable(OdfRect.FromCentimeters(1, 1, 10, 4), 1, 3);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToPowerPointPresentationResult(new PowerPointOpenDocumentConversionOptions {
                MaxTableRows = 2,
                MaxTableColumns = 2
            }));

        Assert.Contains("columns (3)", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointToOdpRejectsTablesBeyondConfiguredBounds() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        source.AddSlide().AddTablePoints(1, 3, 10, 10, 120, 40);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToOpenDocumentResult(new PowerPointOpenDocumentConversionOptions {
                MaxTableRows = 2,
                MaxTableColumns = 2
            }));

        Assert.Contains("columns (3)", exception.Message, StringComparison.Ordinal);
    }
}
