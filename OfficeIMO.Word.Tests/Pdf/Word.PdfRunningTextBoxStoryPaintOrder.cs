using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ReverseDepthRunningBoxesAcrossParagraphsReportBoundedFlow(bool footer) {
        using WordDocument document = CreateRunningStoryPaintOrderControl(10U, 1U, footer: footer);
        string[] before = AnchoredRunningStoryXml(document);

        // Identical geometry with compatible depths must be admitted. In
        // particular, the footer must not fail an unrelated body-paint guard.
        using WordDocument ordered = CreateRunningStoryPaintOrderControl(1U, 10U, footer: footer);
        var orderedResult = ordered.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(orderedResult.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var orderedPdf = PdfPigDocument.Open(orderedResult.Value.ToBytes());
        Assert.DoesNotContain(orderedResult.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        var orderedPage = orderedPdf.GetPage(2);
        var orderedFirst = AnchoredRunningLetters(orderedPage, "FIRSTBOX");
        var orderedSecond = AnchoredRunningLetters(orderedPage, "SECONDBOX");
        if (footer) Assert.Contains("FOOTER COUNT 2", orderedPage.Text);
        // The footer counter-control establishes admission and retained content;
        // its native coordinates are outside this paint-order contract.
        if (!footer) foreach (var letters in new[] { orderedFirst, orderedSecond })
            Assert.Equal((612D - 468D * .585D) / 2D + 7.2D, letters[0].StartBaseLine.X, 3);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string marker in new[] { "PREFIX", "FIRSTBOX", "SUFFIX", "NEXT", "SECONDBOX", "END", "S2" })
            AnchoredRunningLetters(pdf.GetPage(2), marker);
        if (footer) Assert.Contains("FOOTER COUNT 2", pdf.GetPage(2).Text);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Theory]
    [InlineData(1U, 10U)]
    [InlineData(10U, 10U)]
    public void SaveAsPdf_CompatibleDepthRunningParagraphsKeepAnchoredPaintOrder(uint firstDepth, uint secondDepth) {
        using WordDocument document = CreateRunningStoryPaintOrderControl(firstDepth, secondDepth);
        string[] before = AnchoredRunningStoryXml(document);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var page = pdf.GetPage(2);
        string visible = string.Concat(page.Letters.Select(letter => letter.Value));
        Assert.True(visible.IndexOf("FIRSTBOX", StringComparison.Ordinal) < visible.IndexOf("SECONDBOX", StringComparison.Ordinal), visible);
        foreach (string marker in new[] { "FIRSTBOX", "SECONDBOX" })
            Assert.Equal((612D - 468D * .585D) / 2D + 7.2D,
                AnchoredRunningLetters(page, marker)[0].StartBaseLine.X, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_ReverseDepthBoxesWithinOneRunningParagraphKeepLocalSort() {
        using WordDocument document = CreateRunningStoryPaintOrderControl(10U, 1U, sameParagraph: true);
        string[] before = AnchoredRunningStoryXml(document);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var page = pdf.GetPage(2);
        string visible = string.Concat(page.Letters.Select(letter => letter.Value));
        Assert.True(visible.IndexOf("SECONDBOX", StringComparison.Ordinal) < visible.IndexOf("FIRSTBOX", StringComparison.Ordinal), visible);
        foreach (string marker in new[] { "FIRSTBOX", "SECONDBOX" }) AnchoredRunningLetters(page, marker);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_HiddenEarlierRunningBoxDoesNotConstrainVisiblePaintOrder() {
        using WordDocument document = CreateRunningStoryPaintOrderControl(10U, 1U);
        DW.Anchor hidden = document.Sections[1].Header!.First!.TextBoxes.First().Anchor!;
        W.Run host = hidden.Ancestors<W.Run>().Single();
        host.RunProperties ??= new W.RunProperties();
        host.RunProperties.Vanish = new W.Vanish();
        string[] before = AnchoredRunningStoryXml(document);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.DoesNotContain("FIRSTBOX", pdf.GetPage(2).Text);
        Assert.Equal((612D - 468D * .585D) / 2D + 7.2D,
            AnchoredRunningLetters(pdf.GetPage(2), "SECONDBOX")[0].StartBaseLine.X, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    private static WordDocument CreateRunningStoryPaintOrderControl(uint firstDepth, uint secondDepth,
        bool sameParagraph = false, bool footer = false) {
        WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Header header = document.Sections[1].Header!.First!._header!;
        W.Paragraph first = header.Elements<W.Paragraph>().Single();
        first.Descendants<W.TextBoxContent>().Single().Descendants<W.Text>().Single().Text = "FIRSTBOX";
        W.Paragraph second = CloneRunningAnchorParagraph(first, "NEXT", "SECONDBOX", "END");
        first.Descendants<DW.Anchor>().Single().RelativeHeight = firstDepth;
        second.Descendants<DW.Anchor>().Single().RelativeHeight = secondDepth;
        if (sameParagraph) {
            var drawing = second.Descendants<DW.Anchor>().Single().Ancestors<DocumentFormat.OpenXml.AlternateContent>().Single();
            drawing.Remove();
            first.Descendants<DW.Anchor>().Single().Ancestors<W.Run>().Single().Append(drawing);
        } else header.Append(second);
        if (footer) {
            WordHeaderFooter footerStory = document.Sections[1].Footer!.First!;
            W.Footer target = footerStory._footer!;
            foreach (var child in target.ChildElements.ToArray()) child.Remove();
            foreach (W.Paragraph source in header.Elements<W.Paragraph>().ToArray()) {
                source.Remove();
                foreach (DW.Anchor anchor in source.Descendants<DW.Anchor>()) {
                    anchor.VerticalPosition!.PositionOffset!.Text = "9271000"; // 730pt on the authored page.
                    anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
                    // This case isolates box stacking, not foreground body overlap.
                    anchor.BehindDoc = true;
                }
                target.Append(source);
            }
            // SECTIONPAGES requires the rich running-story route whose paragraph
            // canvases and story-peer admission are under test. A plain text-box
            // footer intentionally uses a different projection.
            footerStory.AddParagraph("FOOTER COUNT ").AddField(WordFieldType.SectionPages);
        }
        return document;
    }
}
