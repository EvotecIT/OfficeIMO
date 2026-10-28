using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfHandoutGeometryTests {
    [Theory]
    [InlineData(400, 300)]
    [InlineData(300, 400)]
    [InlineData(1600, 400)]
    public void ThreeSlideHandoutsPreserveAspectRatioAndTheNotesColumn(double slideWidth, double slideHeight) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(slideWidth),OfficeIMO.OfficeLength.Points(slideHeight));
        for (int index = 0; index < 3; index++) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddImage(new MemoryStream(PdfPngTestImages.CreateRgbPng(210, 230, 245)),
                OfficeImageFormat.Png, OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(slideWidth), OfficeIMO.OfficeLength.Points(slideHeight));
            slide.AddTextBox("SLIDE " + (index + 1),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(40));
            PowerPointAutoShape circle = slide.AddShapePoints(OfficePresetShapeType.Ellipse, 20, 100, 100, 100);
            circle.FillColor = "E68A2E"; circle.OutlineColor = "000000";
            slide.Notes.Text = "PRESENTER " + (index + 1);
        }
        presentation.Slides[0].TextBoxes.First().Paragraphs[0].Runs[0].SetHyperlink(presentation.Slides[2]);
        byte[] bytes = presentation.ToPdf(new PowerPointPdfExportOptions {
            PageLayout = PowerPointPdfPageLayout.Handouts, HandoutSlidesPerPage = 3, IncludeSpeakerNotes = true
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        PdfCore.PdfReadPage page = Assert.Single(PdfCore.PdfReadDocument.Open(bytes).Pages);
        (double pageWidth, double pageHeight) = page.GetPageSize();
        var placements = page.GetImagePlacements();
        Assert.Equal(3, placements.Count);
        const double margin = 36D, gutter = 18D;
        double columnWidth = (pageWidth - margin * 2D - gutter) * 0.46D;
        double rowHeight = (pageHeight - margin * 2D - gutter * 2D) / 3D;
        double height = Math.Min(rowHeight - 18D, columnWidth * slideHeight / slideWidth);
        double width = height * slideWidth / slideHeight;
        double left = margin + (columnWidth - width) / 2D;
        for (int index = 0; index < placements.Count; index++) {
            var image = placements[index];
            Assert.InRange(Math.Abs(image.Width / image.Height - slideWidth / slideHeight), 0, 0.0001D);
            Assert.InRange(Math.Abs(image.Width - width), 0, 0.002D);
            Assert.InRange(Math.Abs(image.Height - height), 0, 0.002D);
            Assert.InRange(Math.Abs(image.X - left), 0, 0.002D);
            Assert.InRange(Math.Abs(image.Y - (pageHeight - margin - index * (rowHeight + gutter) - height)), 0, 0.002D);
        }
        var spans = page.GetTextSpans();
        for (int index = 0; index < 3; index++) {
            var title = Assert.Single(spans, span => span.Text == "SLIDE " + (index + 1));
            double top = margin + index * (rowHeight + gutter);
            Assert.InRange(title.X, left, left + width);
            Assert.InRange(title.Y, pageHeight - top - height, pageHeight - top);
            Assert.InRange(title.X + title.Advance, left, left + width);
        }
        double notesLeft = margin + columnWidth + gutter;
        var notes = spans.Where(span => span.Text.StartsWith("PRESENTER", StringComparison.Ordinal)).ToArray();
        Assert.Equal(3, notes.Length);
        Assert.All(notes, span => Assert.InRange(Math.Abs(span.X - notesLeft), 0, 0.002D));
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfNamedDestination target = Assert.Single(info.NamedDestinations, d => d.Name == "powerpoint-slide-3");
        Assert.Equal(1, target.PageNumber);
        Assert.InRange(Math.Abs(target.DestinationLeft!.Value - left), 0, 0.002D);
        Assert.InRange(Math.Abs(target.DestinationTop!.Value - (pageHeight - margin - 2 * (rowHeight + gutter))), 0, 0.002D);
        Assert.Equal("powerpoint-slide-3", Assert.Single(info.LinkAnnotations).DestinationName);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, $"handout-aspect-{slideWidth}-{slideHeight}.pdf"), bytes);
        }
    }
}
