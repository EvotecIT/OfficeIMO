using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfParagraphContextTests {
    [Fact]
    public void CellParagraphsRetainInheritedSpacingIndentsAndIndependentAlignment() {
        using PdfPigDocument expected = PdfPigDocument.Open((CreateCellParagraphs(0).ToPdf(PortableOptions())).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open((CreateCellParagraphs(240).ToPdf(PortableOptions())).ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        var letters = actual.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var second = Assert.Single(letters, letter => letter.Value == "B");
        var third = Assert.Single(letters, letter => letter.Value == "C");
        var unspacedThird = Assert.Single(expected.GetPage(1).Letters, letter => letter.Value == "C");

        Assert.Equal(18D, first.StartBaseLine.Y - second.StartBaseLine.Y, 3);
        Assert.Equal(6D, first.StartBaseLine.X - second.StartBaseLine.X, 3);
        Assert.Equal(12D, first.FontSize, 3);
        Assert.Equal(24D, second.FontSize, 3);
        Assert.True(third.StartBaseLine.X > second.StartBaseLine.X + 200D);
        Assert.Equal(12D, unspacedThird.StartBaseLine.Y - third.StartBaseLine.Y, 3);
        Assert.Equal(
            Assert.Single(expected.GetPage(1).Letters, letter => letter.Value == "A").StartBaseLine.Y,
            first.StartBaseLine.Y, 3);
    }

    [Fact]
    public void EachCellParagraphUsesItsOwnResolvedTabStops() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle inherited = document.AddStyle(1, "Right tab");
        inherited.AddTabStop(2160, RtfTabAlignment.Right);
        RtfTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.AddParagraph("L\tR").SetStyle(1);
        cell.AddParagraph("U\tV").AddTabStop(1440);

        using PdfPigDocument pdf = PdfPigDocument.Open((document.ToPdf(PortableOptions())).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        var left = Assert.Single(letters, letter => letter.Value == "L");
        var right = Assert.Single(letters, letter => letter.Value == "R");
        var lowerLeft = Assert.Single(letters, letter => letter.Value == "U");
        var lowerRight = Assert.Single(letters, letter => letter.Value == "V");

        Assert.Equal(left.StartBaseLine.Y, right.StartBaseLine.Y, 3);
        Assert.Equal(lowerLeft.StartBaseLine.Y, lowerRight.StartBaseLine.Y, 3);
        Assert.Equal(108D, right.EndBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(72D, lowerRight.StartBaseLine.X - lowerLeft.StartBaseLine.X, 3);
        Assert.True(left.StartBaseLine.Y > lowerLeft.StartBaseLine.Y);
    }

    [Theory]
    [InlineData(RtfNoteKind.Footnote, "Footnote 7:")]
    [InlineData(RtfNoteKind.Endnote, "Endnote 7:")]
    [InlineData(RtfNoteKind.Annotation, "Annotation 7 (Alice):")]
    public void AppendedNoteParagraphsRetainLayoutAndEmitTheirLabelOnce(RtfNoteKind kind, string label) {
        RtfDocument unspaced = CreateNoteParagraphs(kind, 0);
        RtfDocument spaced = CreateNoteParagraphs(kind, 240);
        byte[] bytes = (spaced.ToPdf(PortableOptions())).ToBytes();
        using PdfPigDocument expected = PdfPigDocument.Open((unspaced.ToPdf(PortableOptions())).ToBytes());
        using PdfPigDocument actual = PdfPigDocument.Open(bytes);
        Assert.Equal(1, actual.NumberOfPages);
        var letters = actual.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "X");
        var second = Assert.Single(letters, letter => letter.Value == "Y");
        var third = Assert.Single(letters, letter => letter.Value == "Z");

        Assert.Equal(18D, first.StartBaseLine.Y - second.StartBaseLine.Y, 3);
        Assert.Equal(12D, first.FontSize, 3);
        Assert.Equal(24D, second.FontSize, 3);
        Assert.True(third.StartBaseLine.X > second.StartBaseLine.X + 200D);
        Assert.Equal(12D,
            Assert.Single(expected.GetPage(1).Letters, letter => letter.Value == "Z").StartBaseLine.Y - third.StartBaseLine.Y,
            3);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Equal(1, CountOccurrences(text, label));
        Assert.Contains("body7", text, StringComparison.Ordinal);
        Assert.True(text.IndexOf("body7", StringComparison.Ordinal) < text.IndexOf(label, StringComparison.Ordinal));
    }

    [Fact]
    public void AppendedNoteParagraphPageBreakKeepsItsLabelOnTheFirstParagraph() {
        RtfDocument document = CreateNoteParagraphs(RtfNoteKind.Endnote, 0);
        RtfNote note = Assert.Single(document.Blocks.OfType<RtfTable>()).Rows[0].Cells[0].Paragraphs[0].Inlines
            .OfType<RtfGeneratedText>().Single().Note!;
        note.Paragraphs[1].PageBreakBefore = true;

        byte[] bytes = (document.ToPdf(PortableOptions())).ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "X");
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "Y");
        Assert.Single(pdf.GetPage(2).Letters, letter => letter.Value == "Z");
        Assert.DoesNotContain(pdf.GetPage(1).Letters, letter => letter.Value == "Z");
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Equal(1, CountOccurrences(text, "Endnote 7:"));
    }

    [Fact]
    public void RichCellParagraphsKeepMergedImagesLinksNotesAndNestedFallbackInSourceOrder() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 3);
        RtfTableCell cell = table.Rows[0].Cells[0];
        cell.HorizontalMerge = cell.VerticalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
        table.Rows[1].Cells[0].VerticalMerge = RtfTableCellMerge.Continue;
        table.Rows[1].Cells[1].HorizontalMerge = table.Rows[1].Cells[1].VerticalMerge = RtfTableCellMerge.Continue;
        RtfParagraph first = cell.AddParagraph("FIRST ");
        first.AddText("LINK").SetHyperlink(new Uri("https://example.test/cell"));
        first.AddFootnote("7", "ONLY-NOTE");
        RtfImage image = first.AddImage(RtfImageFormat.Png, OfficePngWriter.Encode(new OfficeRasterImage(12, 8, OfficeColor.Red)));
        image.DesiredWidthTwips = 240;
        image.DesiredHeightTwips = 160;
        cell.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("NESTED");
        cell.AddParagraph("LAST").SetAlignment(RtfTextAlignment.Right);
        table.Rows[0].Cells[2].AddParagraph("TOP");
        table.Rows[1].Cells[2].AddParagraph("BOTTOM");

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(PortableOptions());
        PdfCore.TableBlock outputTable = Assert.IsType<PdfCore.TableBlock>(result.Value.Blocks.First());
        Assert.Equal(2, outputTable.Cells[0][0].ColumnSpan);
        Assert.Equal(2, outputTable.Cells[0][0].RowSpan);
        Assert.Single(outputTable.Cells[0][0].Images);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "NestedTableFlattened");
        Assert.Equal(RtfConversionAction.Flattened.ToString(), warning.Details["RtfAction"]);

        byte[] bytes = result.Value.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).GetImages());
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Equal(1, CountOccurrences(text, "FIRST"));
        Assert.Equal(1, CountOccurrences(text, "NESTED"));
        Assert.Equal(1, CountOccurrences(text, "LAST"));
        Assert.True(text.IndexOf("FIRST", StringComparison.Ordinal) < text.IndexOf("NESTED", StringComparison.Ordinal));
        Assert.True(text.IndexOf("NESTED", StringComparison.Ordinal) < text.IndexOf("LAST", StringComparison.Ordinal));
        Assert.Equal(1, CountOccurrences(text, "Footnote 7:"));
        Assert.Equal(1, CountOccurrences(text, "ONLY-NOTE"));
        Assert.Contains("TOP", text, StringComparison.Ordinal);
        Assert.Contains("BOTTOM", text, StringComparison.Ordinal);
        PdfCore.PdfLinkAnnotation link = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).GetLinkAnnotationsByUri("https://example.test/cell"));
        Assert.True(link.Width > 0D && link.Width < 100D);
        Assert.True(link.Height > 0D);
    }

    private static RtfDocument CreateCellParagraphs(int secondSpacingBefore) {
        RtfDocument document = RtfDocument.Create();
        AddExactStyle(document);
        RtfTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        RtfParagraph first = cell.AddParagraph().SetStyle(2).SetIndentation(leftTwips: 360, firstLineTwips: 120);
        AddMixedFontLines(first);
        cell.AddParagraph("C").SetAlignment(RtfTextAlignment.Right)
            .SetIndentation(rightTwips: 360).SetParagraphSpacing(beforeTwips: secondSpacingBefore);
        return document;
    }

    private static RtfDocument CreateNoteParagraphs(RtfNoteKind kind, int secondSpacingBefore) {
        RtfDocument document = RtfDocument.Create();
        AddExactStyle(document);
        var note = new RtfNote(kind) { Author = "Alice" };
        RtfParagraph first = note.AddParagraph().SetStyle(2).SetIndentation(leftTwips: 360);
        first.AddText("X").SetFontSize(12);
        first.AddLineBreak();
        first.AddText("Y").SetFontSize(24);
        note.AddParagraph("Z").SetAlignment(RtfTextAlignment.Right)
            .SetIndentation(rightTwips: 360).SetParagraphSpacing(beforeTwips: secondSpacingBefore);
        document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("body").AddNoteReference(note, "7");
        return document;
    }

    private static void AddExactStyle(RtfDocument document) {
        RtfStyle style = document.AddStyle(1, "Exact spacing");
        style.LineSpacingTwips = -360;
        style.LineSpacingMultiple = false;
        document.AddStyle(2, "Inherited exact spacing").BasedOnStyleId = style.Id;
    }

    private static void AddMixedFontLines(RtfParagraph paragraph) {
        paragraph.AddText("A").SetFontSize(12);
        paragraph.AddLineBreak();
        paragraph.AddText("B").SetFontSize(24);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int offset = 0;
        while ((offset = text.IndexOf(value, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += value.Length;
        }
        return count;
    }

    private static RtfToPdfOptions PortableOptions() => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        PdfOptions = new PdfCore.PdfCreateOptions { DefaultFont = PdfCore.PdfStandardFont.Helvetica, DefaultFontSize = 12 }
    };
}
