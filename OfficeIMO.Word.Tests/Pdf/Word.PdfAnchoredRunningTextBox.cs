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
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));

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
    public void SaveAsPdf_InheritedSectionPagesPreservesOverlappingPageAnchoredPeerTextBoxes() {
        using WordDocument document = CreateOverlappingPageAnchoredPeerControl();
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal(controlPdf.NumberOfPages, pdf.NumberOfPages);
        Page page = pdf.GetPage(2);
        Page controlPage = controlPdf.GetPage(2);
        foreach (string token in new[] { "PREFIX", "SUFFIX" })
            Assert.Equal(AnchoredRunningLetters(controlPage, token).Select(letter => letter.StartBaseLine),
                AnchoredRunningLetters(page, token).Select(letter => letter.StartBaseLine));
        Assert.Equal(AnchoredRunningLetters(controlPage, "S2")[0].StartBaseLine.Y,
            AnchoredRunningLetters(page, "S2")[0].StartBaseLine.Y, 3);

        string visible = string.Concat(page.Letters.Select(letter => letter.Value));
        int peerIndex = visible.IndexOf("PEERBOX", StringComparison.Ordinal);
        Assert.True(peerIndex >= 0, visible);
        Assert.Equal(peerIndex, visible.LastIndexOf("PEERBOX", StringComparison.Ordinal));
        int boxIndex = visible.IndexOf("BOX", peerIndex + "PEERBOX".Length, StringComparison.Ordinal);
        Assert.True(boxIndex > peerIndex, $"Overlapping peer paint order changed: {visible}");
        Assert.Equal(boxIndex, visible.LastIndexOf("BOX", StringComparison.Ordinal));
        Assert.Equal(2, visible.Split(new[] { "BOX" }, StringSplitOptions.None).Length - 1);
        foreach (Letter[] box in new[] {
            page.Letters.Skip(peerIndex).Take("PEERBOX".Length).ToArray(),
            page.Letters.Skip(boxIndex).Take("BOX".Length).ToArray()
        }) {
            Assert.Equal((612D - 468D * 0.585D) / 2D + 7.2D, box[0].StartBaseLine.X, 3);
            Assert.Equal(36D - (14.4D + 3.6D), box[0].StartBaseLine.Y -
                AnchoredRunningLetters(controlPage, "PREFIX")[0].StartBaseLine.Y, 3);
        }
        Assert.DoesNotContain(page.Paths, path => path.IsStroked || path.IsFilled);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_PageAnchoredPeersWithDisallowedOverlapRetainWarnedContentFlow() {
        using WordDocument document = CreateOverlappingPageAnchoredPeerControl();
        DW.Anchor[] anchors = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).ToArray();
        anchors[0].AllowOverlap = true;
        anchors[1].AllowOverlap = false;
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        string visible = string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value));
        int prefixIndex = visible.IndexOf("PREFIX", StringComparison.Ordinal);
        int peerIndex = visible.IndexOf("PEERBOX", StringComparison.Ordinal);
        Assert.True(prefixIndex >= 0 && peerIndex > prefixIndex, visible);
        int boxIndex = visible.IndexOf("BOX", peerIndex + "PEERBOX".Length, StringComparison.Ordinal);
        int suffixIndex = visible.IndexOf("SUFFIX", StringComparison.Ordinal);
        Assert.True(boxIndex > peerIndex && suffixIndex > boxIndex, visible);
        foreach (string marker in new[] { "PREFIX", "PEERBOX", "SUFFIX" })
            Assert.Equal(1, visible.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        Assert.Equal(boxIndex, visible.LastIndexOf("BOX", StringComparison.Ordinal));
        Assert.Equal(2, visible.Split(new[] { "BOX" }, StringSplitOptions.None).Length - 1);
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
            paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(24D);
            WordParagraph next = story.AddParagraph("NEXT");
            next.SetStyleId("Normal");
            next.ContextualSpacing = true;
            next.LineSpacingBefore = next.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        }
        string[] before = AnchoredRunningStoryXml(document);
        string[] controlBefore = AnchoredRunningStoryXml(control);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        // Foreground fill and outline must clear the body, including stroke ink.
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
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
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(2);
        const double width = 468D * 0.585D;
        const double height = 648D * 0.08D;
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
            WordTextBox unsupportedPeer = suffix.AddTextBox("PEERBOX", WordImageTextWrapping.Square);
            unsupportedPeer.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
            suffix.AddText("ENDTEXT");
        }
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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

    private static WordDocument CreateOverlappingPageAnchoredPeerControl() {
        WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph paragraph = document.Sections[1].Header!.First!.ChildElements.OfType<W.Paragraph>().Single();
        DW.Anchor original = paragraph.Descendants<DW.Anchor>().Single();
        original.GetFirstChild<DW.DocProperties>()!.Id = 1U;
        var peer = original.Ancestors<DocumentFormat.OpenXml.AlternateContent>().Single().CloneNode(true);
        peer.Descendants<DW.DocProperties>().Single().Id = 2U;
        peer.Descendants<W.Text>().Single().Text = "PEERBOX";
        // The retained input places the cloned peer in the PREFIX run before
        // the original BOX. Equal-depth paints keep that source order.
        paragraph.Elements<W.Run>().First().Append(peer);
        return document;
    }

    private static WordDocument CreateAnchoredRunningTextBoxControl(bool includeTextBox) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
        source.Sections[0].Margins.Top = source.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Points(72D);
        source.Sections[0].Margins.Left = source.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Points(72D);
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(36D);
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
