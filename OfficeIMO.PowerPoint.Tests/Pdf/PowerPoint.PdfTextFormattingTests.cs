using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfTextFormattingTests {
    [Fact]
    public void NativeRunFormattingProjectsToTypedPdfRuns() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(
            new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun authored = presentation.AddSlide()
            .AddTextBoxPoints("Styled", 10, 10, 200, 40)
            .Paragraphs[0].Runs[0];
        authored.Bold = true;
        authored.Italic = true;
        authored.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
        authored.StrikeStyle = PowerPointStrikeStyle.Double;
        authored.SetSubscript();
        authored.Capitalization = PowerPointCapitalization.SmallCaps;
        authored.Color = "336699";
        authored.FontName = "Aptos";
        authored.FontSizePoints = 14D;

        PdfCore.PdfDocumentConversionResult conversion = presentation.ToPdfDocumentResult();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(conversion.Value.Blocks));
        PdfCore.PdfTextRun run = Assert.Single(
            canvas.Items.OfType<PdfCore.PdfCanvasTextBoxItem>().SelectMany(item => item.Runs),
            item => item.Text == "STYLED");

        Assert.Equal("STYLED", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, run.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikeStyle);
        Assert.Equal(PdfCore.PdfTextBaseline.Subscript, run.Baseline);
        Assert.Equal(PdfCore.PdfColor.FromRgb(51, 102, 153), run.Color);
        Assert.Equal("Aptos", run.FontFamily);
        Assert.Equal(14D, run.FontSize);
        Assert.Contains(conversion.Warnings, warning => warning.Code == "small-caps-approximation");
    }
}
