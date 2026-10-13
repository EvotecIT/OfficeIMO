using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_InheritedSectionPagesPreservesPageAnchoredFirstHeaderTextBox() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));

        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal(controlPdf.NumberOfPages, pdf.NumberOfPages);
        Page page = pdf.GetPage(2);
        Page controlPage = controlPdf.GetPage(2);
        Letter[] prefix = AnchoredRunningLetters(page, "PREFIX");
        Letter[] suffix = AnchoredRunningLetters(page, "SUFFIX");
        Letter[] box = AnchoredRunningLetters(page, "BOX");
        Letter[] controlPrefix = AnchoredRunningLetters(controlPage, "PREFIX");
        Letter[] controlSuffix = AnchoredRunningLetters(controlPage, "SUFFIX");

        Assert.Equal(controlPrefix.Select(letter => letter.StartBaseLine), prefix.Select(letter => letter.StartBaseLine));
        Assert.Equal(controlSuffix.Select(letter => letter.StartBaseLine), suffix.Select(letter => letter.StartBaseLine));
        Assert.Equal(prefix[0].StartBaseLine.Y, suffix[0].StartBaseLine.Y, 3);
        Assert.Equal(AnchoredRunningLetters(controlPage, "S2")[0].StartBaseLine.Y,
            AnchoredRunningLetters(page, "S2")[0].StartBaseLine.Y, 3);

        // Authored page-center anchor, 58.5% of the 468pt margin width,
        // with 7.2pt left inset. Read the raw percentage without truncation.
        Assert.Equal((612D - 468D * 0.585D) / 2D + 7.2D, box[0].StartBaseLine.X, 3);
        // Both stories use the same font. Subtract their authored line origins
        // to avoid binding the contract to a machine-specific font ascent.
        Assert.Equal(36D - (14.4D + 3.6D), box[0].StartBaseLine.Y - controlPrefix[0].StartBaseLine.Y, 3);
        Assert.DoesNotContain(page.Paths, path => path.IsStroked || path.IsFilled);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_AnchoredRunningTextBoxPreservesContextualAfterSpacing() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        foreach (WordDocument source in new[] { document, control }) {
            WordHeaderFooter story = source.Sections[1].Header!.First!;
            WordParagraph paragraph = story.Paragraphs.First();
            paragraph.SetStyleId("Normal");
            paragraph.ContextualSpacing = true;
            paragraph.LineSpacingAfterPoints = 24D;
            WordParagraph next = story.AddParagraph("NEXT");
            next.SetStyleId("Normal");
            next.ContextualSpacing = true;
            next.LineSpacingBeforePoints = next.LineSpacingAfterPoints = 0D;
        }
        string[] before = AnchoredRunningStoryXml(document);
        string[] controlBefore = AnchoredRunningStoryXml(control);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        using var controlPdf = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(controlPdf.NumberOfPages, pdf.NumberOfPages);
        Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), "NEXT")[0].StartBaseLine.Y,
            AnchoredRunningLetters(pdf.GetPage(2), "NEXT")[0].StartBaseLine.Y, 3);
        Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), "S2")[0].StartBaseLine.Y,
            AnchoredRunningLetters(pdf.GetPage(2), "S2")[0].StartBaseLine.Y, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(controlBefore, AnchoredRunningStoryXml(control));
    }

    [Fact]
    public void SaveAsPdf_AnchoredRunningTextBoxPreservesAuthoredFillOutlineAndInsets() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }), true);
        shape.AddChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) {
            Width = 25400
        }, true);
        Wps.TextBodyProperties body = anchor.Descendants<Wps.TextBodyProperties>().Single();
        body.LeftInset = 254000;
        body.TopInset = 127000;
        body.RightInset = 152400;
        body.BottomInset = 76200;
        string[] before = AnchoredRunningStoryXml(document);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        using var controlPdf = PdfPigDocument.Open(control.ToPdfBytes(EmptyRunningOptions()));
        Page page = pdf.GetPage(2);
        const double width = 468D * 0.585D;
        const double height = 648D * 0.2D;
        const double left = (612D - width) / 2D;
        var frames = page.Paths.Where(path => path.GetBoundingRectangle() is { } bounds &&
            Math.Abs(bounds.Left - left) < 0.001D && Math.Abs(bounds.Top - (792D - 14.4D)) < 0.001D &&
            Math.Abs(bounds.Width - width) < 0.001D && Math.Abs(bounds.Height - height) < 0.001D).ToArray();
        var outline = Assert.Single(frames, path => path.IsStroked);
        Assert.Equal(2D, outline.LineWidth, 3);
        Assert.Equal((1D, 0D, 0D), outline.StrokeColor.ToRGBValues());
        Assert.Contains(frames, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D));
        Letter[] box = AnchoredRunningLetters(page, "BOX");
        Assert.Equal(left + 20D, box[0].StartBaseLine.X, 3);
        Assert.Equal(36D - (14.4D + 10D), box[0].StartBaseLine.Y -
            AnchoredRunningLetters(controlPdf.GetPage(2), "PREFIX")[0].StartBaseLine.Y, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_InheritedSectionPagesPreservesUnsupportedPeerAndNestedTextBoxContent(bool nested) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordHeaderFooter story = document.Sections[1].Header!.First!;
        WordParagraph first = story.TextBoxes.Single().Paragraphs.Single();
        first.Text = "FIRSTBOX";
        if (nested) {
            first.AddTextBox("INNERBOX", WordImageTextWrapping.Square);
            first.AddText("AFTERINNER");
        } else {
            WordParagraph suffix = story.Paragraphs.Last();
            suffix.AddTextBox("PEERBOX", WordImageTextWrapping.Square);
            suffix.AddText("ENDTEXT");
        }
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        string visible = string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value));
        string[] markers = nested
            ? new[] { "PREFIX", "FIRSTBOX", "INNERBOX", "AFTERINNER", "SUFFIX" }
            : new[] { "PREFIX", "FIRSTBOX", "SUFFIX", "PEERBOX", "ENDTEXT" };
        int previous = -1;
        foreach (string marker in markers) {
            int position = visible.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(position > previous, $"Missing or reordered {marker}: {visible}");
            Assert.Equal(position, visible.LastIndexOf(marker, StringComparison.Ordinal));
            previous = position;
        }
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    private static WordDocument CreateAnchoredRunningTextBoxControl(bool includeTextBox) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        source.Sections[0].Margins.Top = source.Sections[0].Margins.Bottom = 1440;
        source.Sections[0].Margins.Left = source.Sections[0].Margins.Right = 1440;
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = 720;
        source.HeaderDefaultOrCreate.AddParagraph("COUNT ").AddField(WordFieldType.SectionPages);
        SetInheritedRunningControlStyle(source.AddParagraph("S1"));
        WordSection later = source.AddSection();
        later.AddHeadersAndFooters();
        later.DifferentFirstPage = true;
        WordHeaderFooter story = later.Header!.First!;
        foreach (var child in story.ChildElements.ToArray()) child.Remove();
        WordParagraph paragraph = story.AddParagraph("PREFIX");
        if (includeTextBox) {
            paragraph.AddText(string.Empty).AddTextBox("BOX", WordImageTextWrapping.Square);
            DW.Anchor anchor = paragraph._paragraph.Descendants<DW.Anchor>().Single();
            anchor.HorizontalPosition!.RelativeFrom = DW.HorizontalRelativePositionValues.Page;
            anchor.HorizontalPosition.HorizontalAlignment!.Text = "center";
            anchor.VerticalPosition!.RelativeFrom = DW.VerticalRelativePositionValues.Page;
            anchor.VerticalPosition.PositionOffset!.Text = "182880";
            W14.RelativeWidth width = anchor.Elements<W14.RelativeWidth>().Single();
            width.ObjectId = W14.SizeRelativeHorizontallyValues.Margin;
            width.PercentageWidth!.Text = "58500";
            Wps.TextBodyProperties body = anchor.Descendants<Wps.TextBodyProperties>().Single();
            body.LeftInset = body.RightInset = 91440;
            body.TopInset = body.BottomInset = 45720;
            anchor.Descendants<Wps.ShapeProperties>().Single().RemoveAllChildren<A.Outline>();
        }
        paragraph.AddText("SUFFIX");
        foreach (var reference in later._sectionProperties.Elements<W.HeaderReference>()
            .Where(reference => reference.Type?.Value == W.HeaderFooterValues.Default).ToArray()) reference.Remove();
        SetInheritedRunningControlStyle(source.AddParagraph("S2"));
        source.Paragraphs.Last().AddBreak(WordBreakType.Page);
        SetInheritedRunningControlStyle(source.AddParagraph("S3"));
        return WordDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static Letter[] AnchoredRunningLetters(Page page, string token) {
        string visible = string.Concat(page.Letters.Select(letter => letter.Value));
        int index = visible.IndexOf(token, StringComparison.Ordinal);
        Assert.True(index >= 0, $"Missing {token}: {visible}");
        Assert.Equal(index, visible.LastIndexOf(token, StringComparison.Ordinal));
        return page.Letters.Skip(index).Take(token.Length).ToArray();
    }

    private static string[] AnchoredRunningStoryXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return new[] { main.Document.OuterXml }
            .Concat(main.HeaderParts.OrderBy(part => part.Uri.ToString()).Select(part => part.Header.OuterXml))
            .Concat(main.FooterParts.OrderBy(part => part.Uri.ToString()).Select(part => part.Footer.OuterXml)).ToArray();
    }
}
