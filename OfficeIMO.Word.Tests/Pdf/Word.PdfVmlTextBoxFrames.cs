using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("default", true, true, 0.75D)]
    [InlineData("unfilled", true, false, 0.75D)]
    [InlineData("unstroked", false, true, 0D)]
    [InlineData("explicit", true, false, 2D)]
    [InlineData("transparent", false, false, 0D)]
    [InlineData("child-weight", true, true, 2D)]
    [InlineData("child-disabled", false, true, 0D)]
    [InlineData("rectangle-default", true, true, 0.75D)]
    public void VmlTextBoxRetainsItsFrameWithoutAnExplicitFillColor(string variant, bool stroked, bool filled, double width) {
        using WordDocument document = WordDocument.Create();
        var shape = new V.Shape(new V.TextBox(new TextBoxContent(new Paragraph(
            new Run(new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" }),
                new Text("Visible textbox")))))) {
            Id = "VmlFrame", Type = "#_x0000_t202",
            Style = "position:absolute;left:72pt;top:72pt;width:360pt;height:120pt"
        };
        if (!filled) shape.Filled = false;
        if (!stroked && variant != "child-disabled") shape.Stroked = false;
        if (variant == "child-disabled") shape.Append(new V.Stroke { On = false });
        if (variant == "child-weight") shape.Append(new V.Stroke { Weight = "2pt" });
        if (variant == "rectangle-default") shape.Type = "#_x0000_t1";
        if (variant == "explicit") {
            shape.StrokeColor = "#FF0000";
            shape.StrokeWeight = "2pt";
        }
        document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(new Paragraph(new Run(new Picture(shape)))));
        document.AddParagraph("Body");
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = pdf.GetPage(1);
        Assert.Contains("Visible textbox", page.Text);
        var frames = page.Paths.Where(path => (path.IsStroked || path.IsFilled) &&
            path.GetBoundingRectangle() is { } bounds &&
            Math.Abs(bounds.Width - 360D) < 0.01D && Math.Abs(bounds.Height - 120D) < 0.01D).ToArray();
        if (!stroked && !filled) {
            Assert.Empty(frames);
            return;
        }
        var frame = Assert.Single(frames);
        Assert.Equal(stroked, frame.IsStroked);
        Assert.Equal(filled, frame.IsFilled);
        if (stroked) Assert.InRange(Math.Abs(frame.LineWidth - width), 0D, 0.01D);
    }
}
