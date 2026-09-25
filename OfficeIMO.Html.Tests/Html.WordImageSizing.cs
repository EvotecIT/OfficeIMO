using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_DimensionlessWideImageFitsThePageAndReportsTheChange() {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1600, 600));
        string html = $"<p>Before</p><img src='data:image/png;base64,{image}' alt='Wide photo'><p>After</p>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.Value;
        WordImage drawing = Assert.Single(document.Images);
        WordSection section = document.Sections[0];
        double contentWidth = (section.PageSettings.Width!.Value - section.Margins.Left - section.Margins.Right) / 15D;

        Assert.Equal(contentWidth, drawing.Width!.Value, precision: 2);
        Assert.Equal(contentWidth * 600D / 1600D, drawing.Height!.Value, precision: 2);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "After");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation
            && diagnostic.Source == "Wide photo");
    }

    [Fact]
    public void HtmlToWord_ReusedImageHonorsEachAuthoredSize() {
        string html = $"""
            <p><img src="data:image/png;base64,{ValidPng}" width="100" height="50" alt="Small"></p>
            <p><img src="data:image/png;base64,{ValidPng}" width="200" height="100" alt="Large"></p>
            """;

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.Value;
        WordImage[] images = document.Images.ToArray();

        Assert.Equal(2, images.Length);
        Assert.Equal(100D, images[0].Width!.Value, precision: 2);
        Assert.Equal(50D, images[0].Height!.Value, precision: 2);
        Assert.Equal(200D, images[1].Width!.Value, precision: 2);
        Assert.Equal(100D, images[1].Height!.Value, precision: 2);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated);
    }
}
