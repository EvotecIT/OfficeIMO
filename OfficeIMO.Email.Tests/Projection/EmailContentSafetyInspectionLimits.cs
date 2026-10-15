using OfficeIMO.Email.Data;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailContentSafetyInspectionLimitTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FindingCeilingPreservesTheExplicitAdvisoryProjection(bool exclude) {
        string html = HiddenParagraphs();
        var source = new EmailDocument(); source.Body.Html = html;
        EmailBodyProjectionResult projection = EmailBodyProjection.Create(source, new EmailBodyProjectionOptions {
            IncludeResources = false, InspectContentSafety = true,
            ConcealedTextPolicy = exclude ? EmailConcealedTextPolicy.ExcludeRemovable : EmailConcealedTextPolicy.Preserve
        });
        Assert.Equal("Unavailable", projection.ContentSafety!.InspectionStatus);
        Assert.True(projection.ContentSafety.FindingLimitMayHaveBeenReached);
        Assert.Equal(exclude, projection.ContentSafety.ConcealedTextOmitted);
        Assert.Equal(!exclude, projection.Html.Contains("hidden"));
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_CONTENT_SAFETY_INCOMPLETE");
        Assert.Equal(html, source.Body.Html);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(512)]
    public void FindingCeilingPreservesMetadataAndDistinguishesTheTwoAdvisorySamples(int maxSamples) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".eml");
        string html = HiddenParagraphs();
        try {
            File.WriteAllText(path, "Content-Type: text/html; charset=utf-8\r\n\r\n" + html);
            using EmailDataOpenResult opened = EmailDataArtifact.Open(path);
            EmailHtmlDataInspectionReport report = EmailHtmlDataInspector.Inspect(opened, new EmailDataInspectionOptions(maxSamples: maxSamples));
            Assert.True(report.Metadata.MessageInspected);
            Assert.Single(report.Metadata.Bodies);
            Assert.Equal("Unavailable", report.BodyContentSafety!.InspectionStatus);
            Assert.True(report.BodyContentSafety.FindingLimitMayHaveBeenReached);
            Assert.Equal(maxSamples == 1 ? "InspectionUnavailable" : "Completed", report.HtmlInspectionStatus);
            Assert.Equal(maxSamples == 1, report.FindingLimitMayHaveBeenReached);
            Assert.Equal(maxSamples == 1 ? "EMAIL_HTML_INSPECTION_UNAVAILABLE" : null, report.HtmlDiagnosticCode);
            Assert.Equal(html, opened.EmailDocument!.Body.Html);
        } finally { File.Delete(path); }
    }

    private static string HiddenParagraphs() => string.Concat(Enumerable.Repeat("<p style='display:none'>hidden</p>", 300));
}
