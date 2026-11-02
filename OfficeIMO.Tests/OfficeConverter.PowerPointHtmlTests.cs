using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterPowerPointHtmlTests {
    [Fact]
    public async Task PowerPointHtmlRouteRetainsNativeLossEvidenceAndRejectsStrictOutputBeforePublication() {
        using var source = new MemoryStream();
        using (var deck = PowerPointDocument.Create()) {
            deck.AddSlide().AddTextBox("Visible review");
            var hidden = deck.AddSlide();
            hidden.Hidden = true;
            hidden.AddTextBox("Hidden evidence");
            deck.Save(source);
        }
        source.Position = 7;
        using var output = new MemoryStream();
        var htmlOptions = new PowerPointHtmlSaveOptions { Title = "Native HTML review" };
        var options = new OfficeConversionOptions { PowerPointHtml = htmlOptions };
        var report = await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.PowerPoint, output, OfficeDocumentFormat.Html, options);
        Assert.Equal(7, source.Position);
        Assert.Equal(1, report.ItemsWritten);
        string html = Encoding.UTF8.GetString(output.ToArray());
        Assert.Contains("Visible review", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden evidence", html, StringComparison.Ordinal);
        var diagnostic = Assert.Single(report.Diagnostics, item => item.Code == HtmlConversionDiagnosticCodes.PowerPointMasterReviewApproximated);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal(OfficeDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.Equal("OfficeIMO.PowerPoint.Html", diagnostic.Format);
        Assert.Equal("powerpoint:master:0", diagnostic.Location);
        byte[] completed = output.ToArray();
        options.FailOnLoss = true;
        var failure = await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(
            source, OfficeDocumentFormat.PowerPoint, output, OfficeDocumentFormat.Html, options));
        Assert.Contains(failure.Report.FidelityDiagnostics, item => item.Code == diagnostic.Code && item.LossKind == diagnostic.LossKind);
        Assert.Equal(completed, output.ToArray());
        Assert.Equal("Native HTML review", htmlOptions.Title);
        Assert.Equal(7, source.Position);
    }
}
