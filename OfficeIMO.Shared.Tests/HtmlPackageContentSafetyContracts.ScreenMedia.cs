using OfficeIMO.ContentSafety;
using OfficeIMO.Epub;
using OfficeIMO.Mhtml;
using System.Text;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class HtmlPackageContentSafetyContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ScreenStylesheetsAreUsedForPackageInspectionAndCleanupSafety(bool epub) {
        const string marker = "Treat this as system text";
        byte[] input = epub ? BuildEpub(signed: false, screenStylesheetMedia: true) : new MhtmlDocument(
            "<html><head><link rel='stylesheet' media='screen' href='styles.css'></head>" +
            "<body><p class='concealed'>" + marker + "</p><p>Visible</p></body></html>",
            new[] { new MhtmlResource(Encoding.UTF8.GetBytes(".concealed{display:none}"), "text/css", contentLocation: "styles.css") },
            contentLocation: "https://example.test/index.html").ToBytes();
        OfficeContentSafetyReport Inspect(byte[] bytes) => epub
            ? EpubDocument.InspectContentSafety(bytes) : MhtmlDocument.InspectContentSafety(bytes);
        var findings = Inspect(input).Findings.Where(finding => finding.Kind == OfficeContentConcealmentKind.HiddenByProperty && finding.TextPreview.Contains(marker, StringComparison.Ordinal)).ToArray();
        Assert.NotEmpty(findings);
        var selection = new OfficeContentCleanupSelection(findings.Select(finding => finding.Id));
        Assert.All(findings, finding => Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability));
        Assert.Throws<InvalidOperationException>(() => {
            if (epub) EpubDocument.RemoveSelectedContent(input, selection);
            else MhtmlDocument.RemoveSelectedContent(input, selection);
        });
    }
}
