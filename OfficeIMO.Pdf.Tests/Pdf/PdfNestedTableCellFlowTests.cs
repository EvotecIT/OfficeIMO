using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfNestedTableCellFlowTests {
    [Fact]
    public void OrderedCellKeepsParagraphsAndNestedGridInSourceOrderAtItsActualWidth() {
        var nestedStyle = Style();
        nestedStyle.PreferredWidthFraction = nestedStyle.MaxWidthFraction = .5D;
        nestedStyle.PreserveWidth = true;
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("InnerLeft"), new PdfTableCell("InnerRight") } }, PdfAlign.Left, nestedStyle);
        PdfTableCell cell = Cell(Paragraph("BeforeGrid"), nested, Paragraph("AfterGrid"));
        var outer = Style();
        outer.MaxWidth = 200D; outer.PreserveWidth = true;
        outer.ColumnWidthPoints = new List<double?> { 200D };
        var document = PdfDocument.Create(Options()).TaggedPdfCatalogMarkers().Table(new[] { new[] { cell } }, style: outer);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        var before = Assert.Single(words, word => word.Text == "BeforeGrid");
        var left = Assert.Single(words, word => word.Text == "InnerLeft");
        var right = Assert.Single(words, word => word.Text == "InnerRight");
        var after = Assert.Single(words, word => word.Text == "AfterGrid");
        Assert.True(before.BoundingBox.Bottom > left.BoundingBox.Top);
        Assert.True(left.BoundingBox.Bottom > after.BoundingBox.Top);
        double expectedInnerWidth = (200D - 2D * outer.CellPaddingX) * .5D;
        Assert.Equal(expectedInnerWidth / 2D, right.BoundingBox.Left - left.BoundingBox.Left, 3);
        Assert.Contains(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - expectedInnerWidth) < .001D && bounds.Value.Height > 0D);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(document.ToBytes()).TaggedContent);
        PdfStructureElementInfo documentElement = Assert.Single(tagged.StructureElements, element => element.StructureType == "Document");
        PdfStructureElementInfo outerTable = Assert.Single(tagged.StructureElements,
            element => element.StructureType == "Table" && element.ParentObjectNumber == documentElement.ObjectNumber);
        PdfStructureElementInfo outerRow = Assert.Single(tagged.StructureElements,
            element => element.StructureType == "TR" && element.ParentObjectNumber == outerTable.ObjectNumber);
        PdfStructureElementInfo outerCell = Assert.Single(tagged.StructureElements,
            element => element.StructureType == "TD" && element.ParentObjectNumber == outerRow.ObjectNumber);
        Assert.Single(tagged.StructureElements,
            element => element.StructureType == "Table" && element.ParentObjectNumber == outerCell.ObjectNumber);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "Note");
    }

    [Fact]
    public void ContinuedOuterRowResumesEachCellOnceAndKeepsFollowingParagraphBelowTheLastNestedRow() {
        var nested = new TableBlock(Enumerable.Range(0, 34).Select(index => new[] {
            new PdfTableCell("Nested" + index.ToString("00"))
        }), PdfAlign.Left, Style());
        var document = PdfDocument.Create(Options(190D)).Table(new[] { new[] {
            Cell(Paragraph("BeforeNested"), nested, Paragraph("AfterNested")), new PdfTableCell("ShortPeer")
        } }, style: Style());
        document.Content.Text("FollowingBody", style: ParagraphStyle());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        var all = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (Page: page.Number, Word: word))).ToArray();
        foreach (string token in Enumerable.Range(0, 34).Select(index => "Nested" + index.ToString("00"))
                     .Concat(new[] { "BeforeNested", "AfterNested", "ShortPeer", "FollowingBody" }))
            Assert.Single(all, item => item.Word.Text == token);
        var last = Assert.Single(all, item => item.Word.Text == "Nested33");
        var after = Assert.Single(all, item => item.Word.Text == "AfterNested");
        var following = Assert.Single(all, item => item.Word.Text == "FollowingBody");
        Assert.True(after.Page > last.Page || after.Page == last.Page && after.Word.BoundingBox.Top < last.Word.BoundingBox.Bottom);
        Assert.True(following.Page > after.Page || following.Page == after.Page && following.Word.BoundingBox.Top < after.Word.BoundingBox.Bottom);
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.BoundingBox.Bottom >= 19.99D));
    }

    [Fact]
    public void ContinuedSiblingCellsUseIndependentNestedTableCursors() {
        TableBlock Nested(string prefix, int count) => new TableBlock(Enumerable.Range(0, count)
            .Select(index => new[] { new PdfTableCell(prefix + index.ToString("00")) }), PdfAlign.Left, Style());
        var document = PdfDocument.Create(Options(190D)).Table(new[] { new[] {
            Cell(Nested("Left", 27)), Cell(Paragraph("RightBefore"), Nested("Right", 14), Paragraph("RightAfter"))
        } }, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var all = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        foreach (string token in Enumerable.Range(0, 27).Select(index => "Left" + index.ToString("00"))
                     .Concat(Enumerable.Range(0, 14).Select(index => "Right" + index.ToString("00")))
                     .Concat(new[] { "RightBefore", "RightAfter" }))
            Assert.Single(all, word => word.Text == token);
        Assert.True(pdf.NumberOfPages > 1);
        Assert.All(pdf.GetPages(), page => Assert.Contains(page.GetWords(), word => word.Text.StartsWith("Left", StringComparison.Ordinal)));
    }

    [Fact]
    public void PageNoteNestedTableContinuesThroughTheSameCellAndTableOwner() {
        var document = PdfDocument.Create(Options(190D));
        var nested = new TableBlock(Enumerable.Range(0, 26).Select(index => new[] {
            new PdfTableCell("NoteRow" + index.ToString("00"))
        }), PdfAlign.Left, Style());
        var note = document.Content.DefineFootnote(body => body.Table(new[] {
            new[] { Cell(Paragraph("NoteStart"), nested, Paragraph("NoteEnd")) }
        }, style: Style()));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: ParagraphStyle());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.True(pdf.NumberOfPages > 1);
        foreach (string token in Enumerable.Range(0, 26).Select(index => "NoteRow" + index.ToString("00"))
                     .Concat(new[] { "Reference", "NoteStart", "NoteEnd" }))
            Assert.Single(words, word => word.Text == token);
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.BoundingBox.Bottom >= 19.99D));
    }

    [Fact]
    public void CellCopyOperationsRetainCapturedNestedStylesAndSourceOrder() {
        var style = Style(); style.MaxWidth = style.PreferredWidth = 90D; style.PreserveWidth = true;
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("Captured") } }, PdfAlign.Left, style);
        PdfTableCell source = Cell(Paragraph("First"), nested, Paragraph("Last"));
        PdfTableCell copy = source.WithNamedDestination("CellTarget").WithNoWrap(false).WithShrinkTextToFit(false).WithViewport(null);
        nested.Style!.MaxWidth = nested.Style.PreferredWidth = 160D;
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { copy } }, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Contains(pdf.GetPage(1).Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 90D) < .001D);
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.True(text.IndexOf("First", StringComparison.Ordinal) < text.IndexOf("Captured", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Captured", StringComparison.Ordinal) < text.IndexOf("Last", StringComparison.Ordinal));
    }

    [Fact]
    public void UnsupportedOrderedCellReferencesAreRejectedBeforeTheirNotesCanBeLost() {
        var document = PdfDocument.Create(Options());
        var note = document.Content.DefineFootnote(body => body.Text("RequiredNote"));
        var reference = new RichParagraphBlock(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }, PdfAlign.Left, null, ParagraphStyle());
        document.Content.Table(new[] { new[] { Cell(reference) } }, style: Style());

        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    private static PdfTableCell Cell(params IPdfBlock[] blocks) =>
        new PdfTableCell(Array.Empty<PdfTextRun>(), null, blocks: blocks);

    private static RichParagraphBlock Paragraph(string text) =>
        new(new[] { PdfTextRun.Normal(text) }, PdfAlign.Left, null, ParagraphStyle());

    private static PdfParagraphStyle ParagraphStyle() => new() { FontSize = 8D, LineHeight = 1D, SpacingBefore = 0D, SpacingAfter = 0D };

    private static PdfTableStyle Style() => new() {
        FontSize = 8D, LineHeight = 1D, CellPaddingX = 2D, CellPaddingY = 2D,
        BorderColor = PdfColor.Black, BorderWidth = .5D, HeaderRowCount = 0,
        AllowRowBreakAcrossPages = true, SpacingBefore = 0D, SpacingAfter = 0D
    };

    private static PdfOptions Options(double height = 400D) => new() {
        PageWidth = 320D, PageHeight = height, MarginLeft = 20D, MarginRight = 20D,
        MarginTop = 20D, MarginBottom = 20D, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 8D
    };
}
