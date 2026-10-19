using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public partial class PowerPointImageExportTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void PowerPointSlide_ImageExportMatchesPdfPageForTextTableThemeAndClippedContent(OfficeImageExportFormat format) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(320, 180);
        presentation.SetThemeColorForAllMasters(PowerPointThemeColor.Accent2, "146C94");
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = "FFF9E8";

        PowerPointTextBox title = slide.AddTextBoxPoints("Canonical slide", 18, 12, 260, 32);
        title.FontSize = 18;
        A.Run run = title.Element.Descendants<A.Run>().First();
        A.RunProperties properties = run.RunProperties ??= new A.RunProperties();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.InsertAt(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent2 }), 0);
        title.Paragraphs[0].AddRun(" bold", run => { run.Bold = true; run.FontSize = 18; });

        PowerPointTable table = slide.AddTablePoints(2, 2, 18, 64, 220, 64);
        table.SetColumnWidthsPoints(110, 110);
        table.SetRowHeightsPoints(32, 32);
        table.GetCell(0, 0).Text = "Region";
        table.GetCell(0, 1).Text = "Revenue";
        table.GetCell(1, 0).Text = "North";
        table.GetCell(1, 1).Text = "42";
        table.GetCell(0, 0).FillColor = "146C94";
        table.GetCell(0, 0).Color = "FFFFFF";

        PowerPointAutoShape edge = slide.AddRectanglePoints(-12, 148, 60, 48);
        edge.FillColor = "D7263D";
        edge.OutlineWidthPoints = 0;

        AssertSharedPdfImageRoute(presentation, format, scale: 1.5D);
        PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
        Assert.Contains("Canonical slide", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
        Assert.Contains("Revenue", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
        Assert.Contains(snapshot.TextSpans, span => span.Text.Contains("Canonical", StringComparison.Ordinal)
            && span.Color == OfficeColor.FromRgb(20, 108, 148));
        OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
        Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
        Assert.Equal((320, 180), (image!.Width, image.Height));
        Assert.Equal(OfficeColor.FromRgb(215, 38, 61), image.GetPixel(4, 170));
        Assert.Equal(OfficeColor.FromRgb(255, 249, 232), image.GetPixel(60, 170));
    }

    private static void AssertSharedPdfImageRoute(PowerPointPresentation presentation, OfficeImageExportFormat format, double scale = 1D) {
        // Compare public routes, without relying on the Drawing scene's element types or text representation.
        PdfReadDocument pdf = PdfReadDocument.Open(presentation.ToPdfBytes(new PowerPointToPdfOptions { IncludeHiddenSlides = true }), new PdfLoadOptions { IncludeArtifactText = true });
        PdfReadPage page = pdf.Pages[0];
        OfficeImageExportResult expected = page.ExportImage(format, new PdfImageExportOptions { Scale = scale });
        OfficeImageExportResult actual = presentation.Slides[0].ExportImage(format, new PowerPointImageExportOptions { Scale = scale });
        Assert.Equal(expected.Format, actual.Format);
        Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
        Assert.Equal(expected.Bytes, actual.Bytes);

        PowerPointSlideVisualSnapshot snapshot = presentation.Slides[0].CreateVisualSnapshot();
        Assert.Equal(page.GetTextSpans().Select(span => (span.Text, span.X, span.Y, span.FontSize, span.Color)),
            snapshot.TextSpans.Select(span => (span.Text, span.X, span.Y, span.FontSize, span.Color)));
    }

    private static OfficeRasterImage ReadSvgOutput(OfficeImageExportResult result) {
        Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
        Assert.True(OfficeSvgDrawingReader.TryRead(result.Bytes, out OfficeDrawing? drawing, out int unsupportedFeatures),
            "Exported SVG must be readable by the shared SVG reader.");
        Assert.Equal(0, unsupportedFeatures);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!, result.Width / drawing!.Width, OfficeColor.White);
        Assert.Equal((result.Width, result.Height), (image.Width, image.Height));
        return image;
    }

}
