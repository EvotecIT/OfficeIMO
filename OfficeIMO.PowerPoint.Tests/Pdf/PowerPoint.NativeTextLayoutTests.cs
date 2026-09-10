using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointNativeTextLayoutTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativePdfAndSvgAgreeOnWrappedAndSpacedTextBaselinesAndRetainLinks(bool managedShaping) {
        var fonts = new OfficeFontFaceCollection().Add("NativeFixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont());
        var profile = new OfficeRenderingProfile("native-text-regression", fonts, managedShaping ? OfficeManagedTextShapingProvider.Instance : null);
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(440, 260);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox plain = slide.AddTextBoxPoints("PLAIN baseline", 20, 15, 175, 44);
        plain.FontName = "NativeFixture";
        plain.FontSize = 13;

        PowerPointTextBox wrapped = slide.AddTextBoxPoints("ALPHA", 20, 68, 175, 90);
        wrapped.FontName = "NativeFixture";
        wrapped.FontSize = 13;
        wrapped.Paragraphs[0].AddRun(" beta gamma delta epsilon zeta eta theta", run => run.Italic = true);

        PowerPointTextBox linked = slide.AddTextBoxPoints("LINKONE LINKTWO LINKTHREE ", 230, 15, 160, 90);
        linked.FontName = "NativeFixture";
        linked.FontSize = 13;
        linked.Paragraphs[0].Runs[0].SetHyperlink("https://officeimo.net/native-layout");
        linked.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
        linked.Paragraphs[0].Alignment = PowerPointTextAlignment.Center;
        linked.Paragraphs[0].AddRun("TAIL", run => run.Bold = true);

        PowerPointTextBox spaced = slide.AddTextBoxPoints("FIRST", 20, 175, 375, 75);
        spaced.FontName = "NativeFixture";
        spaced.FontSize = 13;
        spaced.Paragraphs[0].Runs[0].Underline = true;
        spaced.Paragraphs[0].SetSpaceAfterPoints(9);
        PowerPointParagraph second = spaced.AddParagraph("SECOND");
        second.Runs[0].Underline = true;
        second.Alignment = PowerPointTextAlignment.Right;

        var images = new PowerPointImageExportOptions();
        images.UseRenderingProfile(profile);
        XDocument svg = XDocument.Parse(Encoding.UTF8.GetString(slide.ExportImage(OfficeImageExportFormat.Svg, images).Bytes));
        byte[] bytes = presentation.ToPdfBytes(new PowerPointToPdfOptions().UseRenderingProfile(profile));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var letters = page.Letters;
        string pdfText = string.Concat(letters.Select(letter => letter.Value));
        foreach (string marker in new[] { "PLAIN", "ALPHA", "LINKONE", "LINKTHREE", "TAIL", "FIRST", "SECOND" }) {
            XElement element = Assert.Single(svg.Descendants(), element => element.Name.LocalName == "text" && element.Value.StartsWith(marker, StringComparison.Ordinal));
            int index = pdfText.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(index >= 0, "PDF lost searchable text: " + marker);
            double x = double.Parse(element.Attribute("x")!.Value, CultureInfo.InvariantCulture);
            double y = double.Parse(element.Attribute("y")!.Value, CultureInfo.InvariantCulture);
            Assert.True(Math.Abs(letters[index].StartBaseLine.X - x) <= 0.1D, $"{marker}: PDF X={letters[index].StartBaseLine.X}; SVG X={x}; {element}");
            Assert.True(Math.Abs((page.Height - letters[index].StartBaseLine.Y) - y) <= 0.1D, $"{marker}: PDF Y={page.Height - letters[index].StartBaseLine.Y}; SVG Y={y}");
        }
        Assert.Contains("LINKTWO", pdfText);
        Assert.Equal(new[] { "https://officeimo.net/native-layout" }, PdfCore.PdfInspector.Inspect(bytes).LinkUris);
    }
}
