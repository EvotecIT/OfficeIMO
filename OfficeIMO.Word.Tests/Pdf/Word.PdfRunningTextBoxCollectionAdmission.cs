using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("direct")]
    [InlineData("character")]
    [InlineData("wrapped")]
    [InlineData("vml")]
    public void SaveAsPdf_HiddenRunningTextBoxDoesNotPaintOrReserveContent(string sourceKind) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        DW.Anchor anchor = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!;
        W.Run host = anchor.Ancestors<W.Run>().Single();
        host.RunProperties ??= new W.RunProperties();
        if (sourceKind == "character") {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style {
                Type = W.StyleValues.Character, StyleId = "HiddenAnchor",
                StyleRunProperties = new W.StyleRunProperties(new W.Vanish())
            });
            host.RunProperties.RunStyle = new W.RunStyle { Val = "HiddenAnchor" };
        } else host.RunProperties.Vanish = new W.Vanish();
        if (sourceKind == "wrapped") {
            W.Paragraph parent = host.Ancestors<W.Paragraph>().Single();
            var wrapper = new W.SdtRun(new W.SdtContentRun());
            parent.InsertBefore(wrapper, host);
            host.Remove();
            wrapper.SdtContentRun!.Append(host);
        } else if (sourceKind == "vml") {
            foreach (var child in host.ChildElements.Where(child => child is not W.RunProperties).ToArray()) child.Remove();
            var shape = new V.Shape { Id = "HiddenVml", Style = "width:100pt;height:30pt" };
            shape.Append(new V.TextBox(new W.TextBoxContent(new W.Paragraph(new W.Run(new W.Text("HIDDENVML"))))));
            host.Append(new W.Picture(shape));
        }
        string[] before = AnchoredRunningStoryXml(document);
        string stylesBefore = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(controlPdf.NumberOfPages, pdf.NumberOfPages);
        Assert.DoesNotContain("BOX", pdf.GetPage(2).Text);
        foreach (string token in new[] { "PREFIX", "SUFFIX", "S2" })
            Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), token).Select(letter => letter.StartBaseLine),
                AnchoredRunningLetters(pdf.GetPage(2), token).Select(letter => letter.StartBaseLine));
        Assert.DoesNotContain(pdf.GetPage(2).Paths, path => path.IsStroked || path.IsFilled);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(stylesBefore, document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_HiddenPeerCannotRejectVisibleRunningAnchor() {
        using WordDocument document = CreateOverlappingPageAnchoredPeerControl();
        DW.Anchor hidden = document.Sections[1].Header!.First!.ChildElements.SelectMany(child => child.Descendants<DW.Anchor>())
            .Single(anchor => anchor.Descendants<W.Text>().Single().Text == "PEERBOX");
        W.Run host = hidden.Ancestors<W.Run>().Single();
        // The clone shares PREFIX's run. Keep that ordinary text visibly separate.
        var clone = new W.Run(new W.RunProperties(new W.Vanish()));
        var drawing = hidden.Ancestors<DocumentFormat.OpenXml.AlternateContent>().Single();
        host.InsertAfterSelf(clone);
        drawing.Remove();
        clone.Append(drawing);
        hidden.AllowOverlap = false;
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.DoesNotContain("PEERBOX", pdf.GetPage(2).Text);
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Equal((612D - 468D * .585D) / 2D + 7.2D,
            AnchoredRunningLetters(pdf.GetPage(2), "BOX")[0].StartBaseLine.X, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_NonOverlapPeersAcrossRunningParagraphsRetainWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Header root = document.Sections[1].Header!.First!._header!;
        W.Paragraph peer = CloneRunningAnchorParagraph(root.Elements<W.Paragraph>().Single(), "NEXT", "PEERBOX", "END");
        peer.Descendants<DW.Anchor>().Single().AllowOverlap = false;
        root.Append(peer);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string visible = pdf.GetPage(2).Text;
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "NEXT", "PEERBOX", "END", "S2" }) Assert.Contains(token, visible);
        Assert.Equal(1, visible.Split(new[] { "PEERBOX" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Theory]
    [InlineData("roundRect")]
    [InlineData("ellipse")]
    public void SaveAsPdf_NonRectangularRunningTextBoxRetainsWarnedContent(string preset) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        DW.Anchor anchor = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!;
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.GetFirstChild<A.PresetGeometry>()!.Preset = preset == "roundRect"
            ? A.ShapeTypeValues.RoundRectangle : A.ShapeTypeValues.Ellipse;
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }), true);
        shape.AddChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) { Width = 25400 }, true);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ObjectOnlyRunningTextBoxReportsBoundedFlowBeforeSerialization(bool nonRectangular) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Header header = document.Sections[1].Header!.First!._header!;
        foreach (W.Text text in header.Descendants<W.Text>()
            .Where(text => !text.Ancestors<W.TextBoxContent>().Any()).ToArray()) text.Remove();
        if (nonRectangular)
            header.Descendants<A.PresetGeometry>().Single().Preset = A.ShapeTypeValues.RoundRectangle;
        Assert.Empty(header.Descendants<W.Text>().Where(text => !text.Ancestors<W.TextBoxContent>().Any()));
        string[] before = AnchoredRunningStoryXml(document);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "BOX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_SingleNonOverlapAnchorsInDistinctInheritedVariantsRemainIndependent() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph first = document.Sections[1].Header!.First!._header!.Elements<W.Paragraph>().Single();
        first.Descendants<DW.Anchor>().Single().AllowOverlap = false;
        W.Paragraph separate = CloneRunningAnchorParagraph(first, "DP", "DEFAULTBOX", "DE");
        document.Sections[0].Header!.Default!._header!.Append(separate);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("BOX", pdf.GetPage(2).Text);
        Assert.DoesNotContain("DEFAULTBOX", pdf.GetPage(2).Text);
        foreach (int page in new[] { 1, 3 }) Assert.Contains("DEFAULTBOX", pdf.GetPage(page).Text);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_HiddenAnchorInInheritedDefaultStoryRetainsVisibleFields() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph first = document.Sections[1].Header!.First!._header!.Elements<W.Paragraph>().Single();
        W.Paragraph hidden = CloneRunningAnchorParagraph(first, "DEFAULTPREFIX", "HIDDENDEFAULT", "DEFAULTEND");
        W.Run host = hidden.Descendants<DW.Anchor>().Single().Ancestors<W.Run>().Single();
        host.RunProperties ??= new W.RunProperties();
        host.RunProperties.Vanish = new W.Vanish();
        document.Sections[0].Header!.Default!._header!.Append(hidden);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        foreach (int page in new[] { 1, 3 }) {
            Assert.DoesNotContain("HIDDENDEFAULT", pdf.GetPage(page).Text);
            Assert.Contains("DEFAULTPREFIX", pdf.GetPage(page).Text);
            Assert.Contains("DEFAULTEND", pdf.GetPage(page).Text);
        }
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    private static W.Paragraph CloneRunningAnchorParagraph(W.Paragraph source, string prefix, string box, string suffix) {
        var clone = (W.Paragraph)source.CloneNode(true);
        clone.Elements<W.Run>().First().Elements<W.Text>().Single().Text = prefix;
        clone.Elements<W.Run>().Last().Elements<W.Text>().Single().Text = suffix;
        clone.Descendants<W.TextBoxContent>().Single().Descendants<W.Text>().Single().Text = box;
        clone.Descendants<DW.DocProperties>().Single().Id = 2U;
        return clone;
    }
}
