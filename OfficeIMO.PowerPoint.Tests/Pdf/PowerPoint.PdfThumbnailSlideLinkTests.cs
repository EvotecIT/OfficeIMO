using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfThumbnailSlideLinkTests {
    [Theory]
    [InlineData(PowerPointPdfPageLayout.NotesPages, 3, false)]
    [InlineData(PowerPointPdfPageLayout.NotesPages, 3, true)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 3, false)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 3, true)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 4, false)]
    [InlineData(PowerPointPdfPageLayout.Handouts, 4, true)]
    public void ThumbnailLinksResolveToVisibleSlideLocationsAndReportUnavailableTargets(
        PowerPointPdfPageLayout layout, int perPage, bool includeHidden) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(400, 300);
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide hidden = presentation.AddSlide(); hidden.Hidden = true;
        PowerPointSlide third = presentation.AddSlide();
        PowerPointSlide fourth = presentation.AddSlide();
        PowerPointSlide fifth = presentation.AddSlide();
        first.AddTextBoxPoints("VISIBLE TARGET", 20, 20, 350, 40)
            .Paragraphs[0].Runs[0].SetHyperlink(third);
        first.AddTextBoxPoints("HIDDEN TARGET", 20, 65, 350, 40)
            .Paragraphs[0].Runs[0].SetHyperlink(hidden);
        first.AddTextBoxPoints("ABSENT TARGET", 20, 110, 350, 40)
            .Paragraphs[0].Runs[0].Hyperlink = new Uri("#slide-99", UriKind.Relative);
        PowerPointTextBox grouped = first.AddTextBoxPoints("GROUP TARGET", 20, 155, 350, 40);
        grouped.Paragraphs[0].Runs[0].SetHyperlink(fourth);
        grouped.Paragraphs[0].AddRun(" HIDDEN").SetHyperlink(hidden);
        grouped.Paragraphs[0].AddRun(" ABSENT").Hyperlink = new Uri("#slide-99", UriKind.Relative);
        first.GroupShapes(new PowerPointShape[] { grouped, first.AddRectanglePoints(20, 200, 350, 2) });
        PowerPointTable table = first.AddTablePoints(1, 2, 20, 210, 350, 70);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.Text = "TABLE TARGET"; cell.Paragraphs[0].Runs[0].SetHyperlink(fifth);
        PowerPointTableCell flowing = table.GetCell(0, 1);
        flowing.Text = "TABLE FLOW"; flowing.Paragraphs[0].Runs[0].SetHyperlink(third);
        flowing.AddParagraph("TABLE HIDDEN").Runs[0].SetHyperlink(hidden);
        flowing.AddParagraph("TABLE EXTERNAL").Runs[0].Hyperlink = new Uri("https://example.com/table");
        fifth.AddTextBoxPoints("EXTERNAL TARGET", 20, 20, 350, 40)
            .Paragraphs[0].Runs[0].Hyperlink = new Uri("https://example.com/slide");
        var options = new PowerPointPdfExportOptions {
            PageLayout = layout, HandoutSlidesPerPage = perPage, IncludeHiddenSlides = includeHidden
        };

        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence(options);
        byte[] bytes = conversion.ToBytes();
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(bytes);
        int[] visibleNumbers = includeHidden ? new[] { 1, 2, 3, 4, 5 } : new[] { 1, 3, 4, 5 };
        Assert.Equal(layout == PowerPointPdfPageLayout.NotesPages ? visibleNumbers.Length
            : (visibleNumbers.Length + perPage - 1) / perPage, pdf.Pages.Count);
        foreach (int number in visibleNumbers) {
            int position = Array.IndexOf(visibleNumbers, number);
            Assert.Equal(layout == PowerPointPdfPageLayout.NotesPages ? position + 1 : position / perPage + 1,
                Assert.Single(pdf.NamedDestinations, d => d.Name == "powerpoint-slide-" + number).PageNumber);
        }
        foreach (int number in new[] { 3, 4, 5 })
            Assert.Contains(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-" + number);
        Assert.Contains(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-5" && link.Contents == "TABLE TARGET");
        Assert.Contains(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-3" && link.Contents == "TABLE FLOW");
        Assert.Equal(includeHidden, pdf.LinkAnnotations.Any(link => link.DestinationName == "powerpoint-slide-2"));
        Assert.DoesNotContain(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-99");
        Assert.All(pdf.LinkAnnotations.Where(link => link.DestinationName != null),
            link => Assert.Single(pdf.NamedDestinations, d => d.Name == link.DestinationName));
        Assert.Contains(pdf.LinkAnnotations, link => link.Uri == "https://example.com/slide");
        Assert.Contains(pdf.LinkAnnotations, link => link.Uri == "https://example.com/table");
        Assert.Contains(conversion.Report.FidelityDiagnostics, d => d.Code == "internal-slide-target-omitted"
            && d.LossKind == OfficeConversionLossKind.Omission);
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, $"thumbnail-links-{layout}-{perPage}-{includeHidden}.pdf"), bytes);
        }
    }
}
