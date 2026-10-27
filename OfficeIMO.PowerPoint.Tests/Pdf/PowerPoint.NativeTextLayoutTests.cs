using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointNativeTextLayoutTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativePdfAndSnapshotAgreeOnWrappedAndSpacedTextBaselinesAndRetainLinks(bool managedShaping) {
        var fonts = new OfficeFontFaceCollection().Add("NativeFixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont());
        var profile = new OfficeRenderingProfile("native-text-regression", fonts, managedShaping ? OfficeManagedTextShapingProvider.Instance : null);
        using var presentation = PowerPointDocument.Create(new MemoryStream());
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
        byte[] svg = slide.ExportImage(OfficeImageExportFormat.Svg, images).Bytes;
        PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot(images);
        byte[] bytes = presentation.ToPdf(new PowerPointPdfExportOptions().UseRenderingProfile(profile)).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var letters = page.Letters;
        string pdfText = string.Concat(letters.Select(letter => letter.Value));
        string snapshotText = string.Concat(snapshot.TextSpans.Select(span => span.Text));
        foreach (string marker in new[] { "PLAIN", "ALPHA", "LINKONE", "LINKTWO", "LINKTHREE", "TAIL", "FIRST", "SECOND" }) {
            Assert.Contains(marker, pdfText, StringComparison.Ordinal);
            Assert.Contains(marker, snapshotText, StringComparison.Ordinal);
        }
        var canonicalPage = PdfCore.PdfReadDocument.Open(bytes).Pages[0];
        // Compare the public PDF span geometry contract; another reader can use a different font-origin convention.
        Assert.Equal(canonicalPage.GetTextSpans().Select(span => (span.Text, Math.Round(span.X, 3), Math.Round(span.Y, 3), Math.Round(span.FontSize, 3), Math.Round(span.Advance, 3), Math.Round(span.RotationDegrees, 3))),
            snapshot.TextSpans.Select(span => (span.Text, Math.Round(span.X, 3), Math.Round(span.Y, 3), Math.Round(span.FontSize, 3), Math.Round(span.Advance, 3), Math.Round(span.RotationDegrees, 3))));
        Assert.Equal("svg", System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(svg)).Root!.Name.LocalName);
        Assert.Equal(new[] { "https://officeimo.net/native-layout" }, PdfCore.PdfInspector.Inspect(bytes).LinkUris);
    }
}
