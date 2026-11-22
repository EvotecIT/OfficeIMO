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
    public void SaveAsPdf_BackgroundRunningTextBoxOverBodyKeepsAnchoredPaint() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        DW.Anchor anchor = document.Sections[1].PageSetup.FirstPageHeader.ChildElements
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
    public void SaveAsPdf_PageAnchoredTextBoxAdmitsFittingMixedRunSizes() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        WordTextBox textBox = document.Sections[1].PageSetup.FirstPageHeader.TextBoxes.Single();
        DW.Anchor anchor = document.Sections[1].PageSetup.FirstPageHeader.ChildElements
            .SelectMany(child => child.Descendants<DW.Anchor>()).Single();
        anchor.RemoveAllChildren<W14.RelativeHeight>();
        textBox.Height = OfficeLength.Emus(571500L); // Authored 45pt frame, not 20% of the margin height.
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

}
