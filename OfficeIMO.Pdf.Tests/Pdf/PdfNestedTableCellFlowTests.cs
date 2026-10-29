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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RelativeNestedTableWidthExcludesItsOwnFrameSpacingAfterScaling(bool preferredWidth) {
        var nestedStyle = Style();
        if (preferredWidth) nestedStyle.PreferredWidthFraction = .5D;
        else nestedStyle.MaxWidthFraction = .5D;
        nestedStyle.FrameWidthAdjustment = -1D;
        nestedStyle.ResolvedWidthAdjustment = -12D;
        nestedStyle.CellSpacing = 6D;
        nestedStyle.PreserveWidth = nestedStyle.AllowWidthBeyondFrame = true;
        nestedStyle.BorderFrame = new PdfTableBorderFrame {
            Spacing = 6D, HorizontalInset = .5D,
            Border = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.Black, Width = 1D }
        };
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("Left"), new PdfTableCell("Right") } },
            PdfAlign.Left, nestedStyle);
        var outer = Style();
        outer.MaxWidth = 216D; outer.CellPaddingX = 9D; outer.PreserveWidth = true;
        outer.ColumnWidthPoints = new List<double?> { 216D };
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { Cell(nested) } }, style: outer);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = pdf.GetPage(1);
        var words = page.GetWords().ToArray();
        var left = Assert.Single(words, word => word.Text == "Left");
        var right = Assert.Single(words, word => word.Text == "Right");
        // The source half-width includes the perimeter's spacing: 98.5pt
        // centerline frame, 86.5pt grid, and equal cells separated by 6pt.
        Assert.Equal(46.25D, right.BoundingBox.Left - left.BoundingBox.Left, 3);
        Assert.Contains(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 98.5D) < .001D);
        Assert.DoesNotContain(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 110.5D) < .001D);
    }

    [Fact]
    public void CapturedRelativeCellWidthRetainsItsReferenceAndGridOffsetsAfterSourceStyleChanges() {
        var nestedStyle = Style();
        nestedStyle.MaxWidthFraction = .5D;
        nestedStyle.FrameWidthAdjustment = 12D;
        nestedStyle.ResolvedWidthAdjustment = -8D;
        nestedStyle.HorizontalOffset = 7D;
        nestedStyle.PreserveWidth = nestedStyle.AllowWidthBeyondFrame = true;
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("CapturedOffsets") } }, PdfAlign.Left, nestedStyle);
        PdfTableCell captured = Cell(nested).WithNamedDestination("CapturedOffsetCell").WithNoWrap(false);
        nestedStyle.MaxWidthFraction = .9D;
        nestedStyle.FrameWidthAdjustment = 100D;
        nestedStyle.ResolvedWidthAdjustment = 0D;
        nested.Style!.MaxWidthFraction = .75D;
        nested.Style.FrameWidthAdjustment = 50D;
        nested.Style.ResolvedWidthAdjustment = 5D;
        nested.Style.HorizontalOffset = 30D;
        var outer = Style();
        outer.MaxWidth = 216D; outer.CellPaddingX = 9D; outer.PreserveWidth = true;
        outer.ColumnWidthPoints = new List<double?> { 216D };
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { captured } }, style: outer);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = pdf.GetPage(1);
        Assert.Single(page.GetWords(), word => word.Text == "CapturedOffsets");
        // The captured containing frame is 198pt; its retained source projection
        // yields 97pt at the retained 7pt placement offset.
        Assert.Contains(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - 97D) < .001D &&
                Math.Abs(bounds.Value.Left - 36D) < .001D);
    }

    [Fact]
    public void RelativeFixedGridRetainsPhysicalColumnMinimumsAfterItsOwnSpacingIsRemoved() {
        const double spacing = 6D, fraction = .1D;
        var minimums = new List<double?> { 20D, 30D };
        double minimumGrid = minimums.Sum(width => width ?? 0D) + spacing;
        double expectedPerimeter = minimumGrid + 2D * spacing;
        var nestedStyle = Style();
        nestedStyle.MaxWidthFraction = fraction;
        nestedStyle.FrameWidthAdjustment = -1D;
        nestedStyle.ResolvedWidthAdjustment = -2D * spacing;
        nestedStyle.CellSpacing = spacing;
        nestedStyle.ColumnWidthPoints = new List<double?> { 45D, 105D };
        nestedStyle.ColumnMinWidthPoints = minimums;
        nestedStyle.CompressFixedColumnsAboveMinimums = true;
        nestedStyle.PreserveWidth = nestedStyle.AllowWidthBeyondFrame = true;
        nestedStyle.BorderFrame = new PdfTableBorderFrame {
            Spacing = spacing, HorizontalInset = .5D,
            Border = new PdfCellBorder { Color = global::OfficeIMO.OfficeColor.Black, Width = 1D }
        };
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("L"), new PdfTableCell("R") } },
            PdfAlign.Left, nestedStyle);
        var outer = Style();
        outer.MaxWidth = 216D; outer.CellPaddingX = 9D; outer.PreserveWidth = true;
        outer.ColumnWidthPoints = new List<double?> { 216D };
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { Cell(nested) } }, style: outer);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("LR", string.Concat(page.Letters.Select(letter => letter.Value)));
        var left = Assert.Single(page.Letters, letter => letter.Value == "L");
        var right = Assert.Single(page.Letters, letter => letter.Value == "R");
        // The parent content frame is198pt; its post-scale target is smaller
        // than the authored20+30pt physical floors plus the6pt grid gap.
        Assert.Equal(minimums[0]!.Value + spacing, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Contains(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Width - expectedPerimeter) < .001D);
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

    [Theory]
    [InlineData(false, 0D)]
    [InlineData(true, 0D)]
    [InlineData(true, 7D)]
    public void FloatingBookmarkFollowsTheFirstTableRowAfterAutomaticPageClearance(bool orderedCells, double spacingBefore) {
        var floating = Style();
        floating.ColumnWidthPoints = new List<double?> { 280D };
        floating.MinRowHeight = 190D;
        floating.Position = new PdfTablePosition();
        var targetStyle = Style();
        targetStyle.SpacingBefore = spacingBefore;
        PdfTableCell target = orderedCells
            ? Cell(Paragraph("TargetFirstLine"), Paragraph("TargetSecondLine"))
            : new PdfTableCell("TargetFirstLine\nTargetSecondLine");
        var document = PdfDocument.Create(Options(240D))
            .Table(new[] { new[] { "FloatingTable" } }, style: floating)
            .Bookmark("ContainedTableStart")
            .Flow(flow => flow.Table(new[] { new[] {
                target
            } }, style: targetStyle), new PdfFlowOptions { KeepTogether = true });
        document.Content.Text("FollowingBody", style: ParagraphStyle());

        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "FloatingTable");
        var first = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "TargetFirstLine");
        var following = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "FollowingBody");
        Assert.True(first.BoundingBox.Bottom > following.BoundingBox.Top);
        var destination = Assert.Single(PdfInspector.Inspect(bytes).NamedDestinations,
            item => item.Name == "ContainedTableStart");
        Assert.Equal(2, destination.PageNumber);
        Assert.Equal(220D - spacingBefore, destination.DestinationTop!.Value, 3);
        Assert.InRange(destination.DestinationTop.Value - first.BoundingBox.Top, 0D, 12D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedHeadersKeepOneLinkPerTargetOnEachPageAndTheFirstNamedDestination(bool orderedCells) {
        const string cellUri = "https://example.com/header-cell", runUri = "https://example.com/header-run";
        var style = Style();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        PdfTableCell cellHeader = orderedCells
            ? new PdfTableCell(Array.Empty<PdfTextRun>(), null, linkUri: cellUri, linkContents: "HeaderCell",
                namedDestinationName: "FirstHeader", blocks: new[] { Paragraph("HeaderCell") })
            : new PdfTableCell("HeaderCell", linkUri: cellUri, namedDestinationName: "FirstHeader");
        var linkedRuns = new[] { PdfTextRun.Link("HeaderRun", runUri) };
        PdfTableCell runHeader = orderedCells
            ? Cell(new RichParagraphBlock(linkedRuns, PdfAlign.Left, null, ParagraphStyle()))
            : new PdfTableCell(linkedRuns);
        var rows = new List<PdfTableCell[]> { new[] { cellHeader, runHeader } };
        rows.AddRange(Enumerable.Range(0, 30).Select(index => new[] {
            new PdfTableCell("Body" + index.ToString("00")), new PdfTableCell("Peer" + index.ToString("00"))
        }));
        var document = PdfDocument.Create(Options(190D)).TaggedPdfCatalogMarkers().Table(rows, style: style);

        byte[] bytes = document.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        var inspection = PdfInspector.Inspect(bytes);
        foreach (var page in pdf.GetPages()) {
            Assert.Single(page.GetWords(), word => word.Text == "HeaderCell");
            Assert.Single(page.GetWords(), word => word.Text == "HeaderRun");
            Assert.Single(inspection.LinkAnnotations, link => link.PageNumber == page.Number && link.Uri == cellUri);
            Assert.Single(inspection.LinkAnnotations, link => link.PageNumber == page.Number && link.Uri == runUri);
        }
        Assert.Equal(2 * pdf.NumberOfPages, inspection.LinkAnnotations.Count);
        var destination = Assert.Single(inspection.NamedDestinations, item => item.Name == "FirstHeader");
        Assert.Equal(1, destination.PageNumber);
        Assert.Equal(170D, destination.DestinationTop!.Value, 3);
        foreach (string token in Enumerable.Range(0, 30).Select(index => "Body" + index.ToString("00")))
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == token);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizedKeptParagraphInAnOrderedRowContinuesWithoutLosingItsPeers(bool noteTable) {
        var document = PdfDocument.Create(Options(190D));
        var expected = Enumerable.Range(0, 40).Select(index => "KeptLine" + index.ToString("00")).ToArray();
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", expected)) };
        var paragraph = new PdfTableCellParagraph(runs, fontSize: 8D, lineSpacing: PdfLineSpacing.Exactly(10D),
            widowControl: false, keepTogether: true);
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("NestedPeer") } }, PdfAlign.Left, Style());
        var rows = new[] { new[] { new PdfTableCell(runs, new[] { paragraph }), Cell(nested), new PdfTableCell("ShortPeer") } };
        if (noteTable) {
            var note = document.Content.DefineFootnote(body => body.Table(rows, style: Style()));
            document.Content.Paragraph(body => body.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: ParagraphStyle());
        } else document.Content.Table(rows, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.Equal(expected, words.Where(word => word.Text.StartsWith("KeptLine", StringComparison.Ordinal)).Select(word => word.Text));
        Assert.Single(words, word => word.Text == "NestedPeer");
        Assert.Single(words, word => word.Text == "ShortPeer");
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.BoundingBox.Bottom >= 19.99D));
    }

    [Fact]
    public void FittingKeptParagraphInAnOrderedRowMovesTogetherWhenTheCurrentFrameIsShort() {
        var document = PdfDocument.Create(Options(190D));
        document.Content.Text("Prefix", style: ParagraphStyle());
        document.Content.Spacer(80D);
        var expected = Enumerable.Range(0, 6).Select(index => "FittingLine" + index.ToString("00")).ToArray();
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", expected)) };
        var paragraph = new PdfTableCellParagraph(runs, fontSize: 8D, lineSpacing: PdfLineSpacing.Exactly(10D),
            widowControl: false, keepTogether: true);
        var nested = new TableBlock(new[] { new[] { new PdfTableCell("FittingPeer") } }, PdfAlign.Left, Style());
        document.Content.Table(new[] { new[] { new PdfTableCell(runs, new[] { paragraph }), Cell(nested) } }, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("FittingLine", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("FittingPeer", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var words = pdf.GetPage(2).GetWords().ToArray();
        Assert.Equal(expected, words.Where(word => word.Text.StartsWith("FittingLine", StringComparison.Ordinal)).Select(word => word.Text));
        Assert.Single(words, word => word.Text == "FittingPeer");
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
        BorderColor = global::OfficeIMO.OfficeColor.Black, BorderWidth = .5D, HeaderRowCount = 0,
        AllowRowBreakAcrossPages = true, SpacingBefore = 0D, SpacingAfter = 0D
    };

    private static PdfCreateOptions Options(double height = 400D) => new() {
        PageWidth = 320D, PageHeight = height, MarginLeft = 20D, MarginRight = 20D,
        MarginTop = 20D, MarginBottom = 20D, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 8D
    };
}
