using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using UglyToad.PdfPig.Content;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_ExactMixedRunningTextBoxWithClippedInkRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        box.Anchor!.RemoveAllChildren<W14.RelativeHeight>();
        box.Height = 571500L;
        box.Anchor.Descendants<A.Transform2D>().Single().Extents!.Cy = 571500L;
        WordParagraph inner = box.Paragraphs.Single();
        foreach (W.Break lineBreak in inner._paragraph.Descendants<W.Break>().ToArray()) lineBreak.Remove();
        inner.Text = "small";
        inner.FontSize = global::OfficeIMO.OfficeLength.Points(6);
        inner.AddText("LARGE").FontSize = global::OfficeIMO.OfficeLength.Points(24);
        inner.LineSpacingPoints = 6;
        inner.LineSpacingRule = WordLineSpacingRule.Exact;
        // Short advances fit; the larger run's actual line ink crosses the inner clip.
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "small", "LARGE", "SUFFIX" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_BackgroundRunningTextBoxOverBodyKeepsAnchoredPaint() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.BehindDoc = true;
        anchor.VerticalPosition!.PositionOffset!.Text = "914400";
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.NoFill>();
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }), true);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), "S2")[0].StartBaseLine,
            AnchoredRunningLetters(pdf.GetPage(2), "S2")[0].StartBaseLine);
        AnchoredRunningLetters(pdf.GetPage(2), "BOX");
        Assert.Contains(pdf.GetPage(2).Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D));
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningTextBoxStrokeOverBodyRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
        anchor.Descendants<Wps.ShapeProperties>().Single().AddChild(
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) { Width = 254000 }, true);
        // The frame ends at 66.24pt; its 20pt outline extends into the 72pt body.
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningTextBoxBelowBodyMarginRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.RemoveAllChildren<W14.RelativeHeight>();
        anchor.Extent!.Cy = 571500L;
        anchor.Descendants<A.Transform2D>().Single().Extents!.Cy = 571500L;
        anchor.VerticalPosition!.PositionOffset!.Text = "9271000"; // 730pt, below the authored 720pt body bottom.
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningTextBoxOverBodyRetainsWarnedContentFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.BehindDoc = false;
        anchor.VerticalPosition!.PositionOffset!.Text = "914400";
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.NoFill>();
        shape.RemoveAllChildren<A.SolidFill>();
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }), true);
        W.Text body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!
            .Descendants<W.Text>().Single(text => text.Text == "S2");
        body.Text += " BODY TEXT INTERSECTS THE FOREGROUND FRAME";
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string marker in new[] { "PREFIX", "BOX", "SUFFIX", "S2" })
            AnchoredRunningLetters(pdf.GetPage(2), marker);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_OversizedMixedRunningTextBoxRetainsWarnedContentFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordParagraph inner = document.Sections[1].Header!.First!.TextBoxes.Single().Paragraphs.Single();
        foreach (W.Break lineBreak in inner._paragraph.Descendants<W.Break>().ToArray()) lineBreak.Remove();
        inner.Text = "x";
        inner.FontFamily = "Arial";
        inner.FontSize = global::OfficeIMO.OfficeLength.Points(6);
        WordParagraph large = inner.AddText("X");
        large.FontFamily = "Arial";
        large.FontSize = global::OfficeIMO.OfficeLength.Points(144);
        large.AddBreak();
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        AnchoredRunningLetters(pdf.GetPage(2), "xX");
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_AsymmetricMirroredRunningTextBoxRetainsWarnedContentFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        document.Settings.MirrorMargins = true;
        WordSection selectingSection = document.Sections[1];
        selectingSection.Margins.Left = global::OfficeIMO.OfficeLength.Points(72D);
        selectingSection.Margins.Right = global::OfficeIMO.OfficeLength.Points(144D);
        // This section's first header is on visible page 2. The unswapped 72pt
        // text origin clears the centered box; the actual 144pt origin does not.
        // Preserve the warned flow until admission uses the effective page frame.
        string[] before = AnchoredRunningStoryXml(document);
        string settingsBefore = document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml;

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string marker in new[] { "PREFIX", "BOX", "SUFFIX" })
            AnchoredRunningLetters(pdf.GetPage(2), marker);
        Assert.Contains("S1", pdf.GetPage(1).Text);
        Assert.Contains("S2", pdf.GetPage(2).Text);
        Assert.Contains("S3", pdf.GetPage(3).Text);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(settingsBefore, document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_PageAnchoredTextBoxAdmitsFittingMixedRunSizes() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        WordTextBox textBox = document.Sections[1].Header!.First!.TextBoxes.Single();
        DW.Anchor anchor = document.Sections[1].Header!.First!.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.RemoveAllChildren<W14.RelativeHeight>();
        textBox.Height = 571500L; // Authored 45pt frame, not 20% of the margin height.
        anchor.Descendants<A.Transform2D>().Single().Extents!.Cy = 571500L;
        WordParagraph inner = textBox.Paragraphs.Single();
        foreach (W.Break lineBreak in inner._paragraph.Descendants<W.Break>().ToArray()) lineBreak.Remove();
        inner.Text = "small";
        inner.FontFamily = "Arial";
        inner.FontSize = global::OfficeIMO.OfficeLength.Points(6);
        WordParagraph large = inner.AddText("LARGE");
        large.FontFamily = "Arial";
        large.FontSize = global::OfficeIMO.OfficeLength.Points(24);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        Page page = pdf.GetPage(2);
        const double frameLeft = (612D - 468D * 0.585D) / 2D;
        const double clipLeft = frameLeft + 7.2D;
        const double clipRight = frameLeft + 468D * 0.585D - 7.2D;
        const double clipTop = 792D - 14.4D - 3.6D;
        const double clipBottom = 792D - 14.4D - 45D + 3.6D;
        foreach (string token in new[] { "small", "LARGE" }) {
            Letter[] letters = AnchoredRunningLetters(page, token);
            Assert.All(letters, letter => {
                Assert.Equal(token == "small" ? 6D : 24D, letter.PointSize, 3);
                Assert.InRange(letter.BoundingBox.Left, clipLeft - 0.001D, clipRight + 0.001D);
                Assert.InRange(letter.BoundingBox.Right, clipLeft - 0.001D, clipRight + 0.001D);
                Assert.InRange(letter.BoundingBox.Bottom, clipBottom - 0.001D, clipTop + 0.001D);
                Assert.InRange(letter.BoundingBox.Top, clipBottom - 0.001D, clipTop + 0.001D);
            });
        }
        foreach (string token in new[] { "PREFIX", "SUFFIX" })
            Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), token).Select(letter => letter.StartBaseLine),
                AnchoredRunningLetters(page, token).Select(letter => letter.StartBaseLine));
        Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), "S2")[0].StartBaseLine.Y,
            AnchoredRunningLetters(page, "S2")[0].StartBaseLine.Y, 3);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Theory]
    [InlineData(WordLineSpacingRule.Exact)]
    [InlineData(WordLineSpacingRule.AtLeast)]
    public void SaveAsPdf_PageAnchoredTextBoxRejectsOversizedAuthoredLineSpacing(WordLineSpacingRule rule) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordParagraph inner = document.Sections[1].Header!.First!.TextBoxes.Single().Paragraphs.Single();
        inner.Text = "SMALL";
        inner.FontFamily = "Arial";
        inner.FontSize = global::OfficeIMO.OfficeLength.Points(6);
        inner.LineSpacingPoints = 144;
        inner.LineSpacingRule = rule;
        // The unchanged 129.6pt frame has 122.4pt after vertical insets.
        // A 144pt authored advance cannot pass canonical canvas line-box fit.
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string marker in new[] { "PREFIX", "SMALL", "SUFFIX" })
            Assert.Equal(1, visible.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        Assert.Contains("COUNT 2", visible);
        Assert.DoesNotContain("{sectionpages}", visible);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }
}
