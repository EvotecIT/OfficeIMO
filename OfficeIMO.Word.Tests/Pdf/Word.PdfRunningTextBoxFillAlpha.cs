using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TranslucentRunningTextBoxOutlineRetainsWarnedContent(bool footer) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        WordHeaderFooter header = document.Sections[1].Header!.First!;
        WordTextBox box = header.TextBoxes.Single();
        Wps.ShapeProperties shape = box.Anchor!.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.Outline>();
        shape.AddChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex(
            new A.Alpha { Val = 65000 }, new A.AlphaOffset { Val = -15000 }) { Val = "FF0000" })) {
            Width = 25400
        }, true);
        if (footer) {
            WordHeaderFooter target = document.Sections[1].Footer!.First!;
            foreach (var child in target.ChildElements.ToArray()) child.Remove();
            foreach (W.Paragraph source in header._header!.Elements<W.Paragraph>().ToArray()) {
                source.Remove();
                target._footer!.Append(source);
            }
            box.Anchor.VerticalPosition!.PositionOffset!.Text = "9271000";
            box.Anchor.BehindDoc = true;
            target.AddParagraph("FOOTER COUNT ").AddField(WordFieldType.SectionPages);
        }
        // The fixed running canvas has no border-alpha channel. It must retain the
        // explicitly approximated route rather than silently paint an opaque border.
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RunningTextBoxFillAlphaReachesHeaderAndFooterCanvasOnce(bool footer) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        WordHeaderFooter header = document.Sections[1].Header!.First!;
        WordTextBox box = header.TextBoxes.Single();
        Wps.ShapeProperties shape = box.Anchor!.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.NoFill>();
        shape.RemoveAllChildren<A.SolidFill>();
        shape.AddChild(new A.SolidFill(new A.SchemeColor(new A.Alpha { Val = 45000 }) {
            Val = A.SchemeColorValues.Accent1
        }), true);
        shape.AddChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) {
            Width = 25400
        }, true);
        if (footer) {
            WordHeaderFooter target = document.Sections[1].Footer!.First!;
            foreach (var child in target.ChildElements.ToArray()) child.Remove();
            foreach (W.Paragraph source in header._header!.Elements<W.Paragraph>().ToArray()) {
                source.Remove();
                target._footer!.Append(source);
            }
            box.Anchor.VerticalPosition!.PositionOffset!.Text = "9271000"; // 730pt on the page.
            box.Anchor.BehindDoc = true;
            target.AddParagraph("FOOTER COUNT ").AddField(WordFieldType.SectionPages);
        }
        string[] before = AnchoredRunningStoryXml(document);

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        byte[] bytes = result.Value.ToBytes();
        using (var pdf = PdfPigDocument.Open(bytes)) {
            Assert.Equal(3, pdf.NumberOfPages);
            foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" })
                AnchoredRunningLetters(pdf.GetPage(2), token);
            Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        }

        var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[1].GetIdentityVisualPrimitives();
        var frame = Assert.Single(primitives, primitive => primitive.FillColor == OfficeColor.FromRgb(68, 114, 196));
        Assert.InRange(frame.FillOpacity ?? 1D, 0.45D - 1D / 255D, 0.45D + 1D / 255D);
        Assert.Equal((612D - 468D * 0.585D) / 2D, frame.X, precision: 3);
        Assert.Equal(footer ? 730D : 14.4D, frame.Y, precision: 3);
        Assert.Equal(468D * 0.585D, frame.Width, precision: 3);
        Assert.Equal(648D * 0.08D, frame.Height, precision: 3);
        Assert.Equal(OfficeColor.Red, frame.StrokeColor);
        Assert.Equal(2D, frame.StrokeWidth, precision: 3);
        Assert.Equal(1D, frame.StrokeOpacity ?? 1D, precision: 6);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }
}
