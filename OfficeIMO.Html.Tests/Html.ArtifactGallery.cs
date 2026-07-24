using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml() {
        const string imageDataUri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII=";
        string html = """
            <!doctype html>
            <html lang="en">
            <head>
                <title>Quarterly Report</title>
                <style>
                    body { font-family: Calibri; }
                    table.report { width: 100%; border-collapse: separate; border-spacing: 6pt; }
                    th, td { border: 1px solid #444; padding: 4pt; vertical-align: middle; }
                    tfoot td { background-color: #eeeeee; font-weight: bold; }
                </style>
            </head>
            <body>
                <article id="report">
                    <h1>Quarterly Report</h1>
                    <p>Prepared for <strong>OfficeIMO</strong> HTML conversion validation.</p>
                    <ul>
                        <li><input type="checkbox" checked> Revenue reviewed</li>
                        <li><input type="checkbox"> Risks pending</li>
                    </ul>
                    <table class="report">
                        <thead>
                            <tr><th>Metric</th><th>Value</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>Revenue</td><td>$42,000</td></tr>
                            <tr><td>Margin</td><td>18%</td></tr>
                        </tbody>
                        <tfoot>
                            <tr><td>Total</td><td>$42,000</td></tr>
                        </tfoot>
                    </table>
                    <label>Owner <input name="owner" value="Ada Lovelace"></label>
                    <label>Status <select name="status"><option>Draft</option><option selected>Approved</option></select></label>
                    <figure><img src="IMAGE_DATA_URI" alt="Roundtrip badge" width="48" height="48"><figcaption>Embedded image proof</figcaption></figure>
                    <!-- skipped comments are diagnostic evidence, not visible document content -->
                </article>
            </body>
            </html>
            """.Replace("IMAGE_DATA_URI", imageDataUri);

        string artifactDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.HtmlArtifactGallery");
        Directory.CreateDirectory(artifactDirectory);
        string inputPath = Path.Combine(artifactDirectory, "quarterly-report.input.html");
        string docxPath = Path.Combine(artifactDirectory, "quarterly-report.docx");
        string roundTripPath = Path.Combine(artifactDirectory, "quarterly-report.roundtrip.html");
        string manifestPath = Path.Combine(artifactDirectory, "quarterly-report.manifest.md");
        string manifestJsonPath = Path.Combine(artifactDirectory, "quarterly-report.manifest.json");

        HtmlCapabilityGalleryManifest manifest = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.Document,
            Trust = HtmlInputTrust.Trusted
        }).SaveHtmlCapabilityGallery(artifactDirectory, new WordHtmlCapabilityGalleryOptions {
            ScenarioId = "quarterly-report",
            Title = "Quarterly Report"
        });
        string roundTripHtml = File.ReadAllText(roundTripPath);
        string manifestMarkdown = File.ReadAllText(manifestPath);
        string manifestJson = File.ReadAllText(manifestJsonPath);

        Assert.Contains("<h1>Quarterly Report</h1>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<thead>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<tfoot>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"checkbox\"", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<select", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<img", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Roundtrip badge", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("skipped comments are diagnostic evidence", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("quarterly-report", manifest.Result.Scenario.Id);
        Assert.Equal(3, manifest.Result.Artifacts.Count);
        Assert.Contains(manifest.Result.Artifacts, artifact => artifact.Kind == "docx" && artifact.Length > 0 && artifact.Sha256.Length == 64);
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Component == "OfficeIMO.Word.Html" && diagnostic.Code == "HtmlCommentSkipped");
        Assert.Contains(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Component == "OfficeIMO.Word.Html" && diagnostic.Code == "WordOpenXmlPackageValid");
        Assert.DoesNotContain(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "WordOpenXmlValidationError");
        Assert.DoesNotContain(manifest.Result.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
        Assert.Equal(1, manifest.ResourceManifest.AllowedCount);
        Assert.Equal(0, manifest.ResourceManifest.BlockedCount);
        Assert.Equal(new[] { OfficeHtmlConversionProfile.WordDocumentRoundTrip }, manifest.OfficeProfiles);
        Assert.True(File.Exists(inputPath));
        Assert.True(File.Exists(docxPath));
        Assert.True(File.Exists(roundTripPath));
        Assert.True(File.Exists(manifestPath));
        Assert.True(File.Exists(manifestJsonPath));
        Assert.Contains("Profile: Document", manifestMarkdown);
        Assert.Contains("Profile Contract", manifestMarkdown);
        Assert.Contains("Office Profile Contracts", manifestMarkdown);
        Assert.Contains("Word Document Roundtrip (Word -> Document)", manifestMarkdown);
        Assert.Contains("Roundtrip Expectations", manifestMarkdown);
        Assert.Contains("Preserved: form controls => roundtrip HTML contains form control elements", manifestMarkdown);
        Assert.Contains("Preserved: images => roundtrip HTML contains image or SVG evidence", manifestMarkdown);
        Assert.Contains("Preserved: docx package => generated DOCX passes OpenXML validation", manifestMarkdown);
        Assert.Contains("HtmlCommentSkipped", manifestMarkdown, StringComparison.OrdinalIgnoreCase);

        using JsonDocument manifestJsonDocument = JsonDocument.Parse(File.ReadAllText(manifestJsonPath));
        JsonElement manifestRoot = manifestJsonDocument.RootElement;
        Assert.Equal("officeimo.html.capability-gallery", manifestRoot.GetProperty("schemaId").GetString());
        Assert.Equal("quarterly-report", manifestRoot.GetProperty("scenario").GetProperty("id").GetString());
        Assert.Equal("Document", manifestRoot.GetProperty("profile").GetProperty("id").GetString());
        JsonElement officeProfiles = manifestRoot.GetProperty("officeProfiles");
        Assert.Equal(1, officeProfiles.GetArrayLength());
        Assert.Equal("WordDocumentRoundTrip", officeProfiles[0].GetProperty("id").GetString());
        Assert.Equal("Word", officeProfiles[0].GetProperty("sourceFormat").GetString());
        Assert.Equal("Document", officeProfiles[0].GetProperty("sharedProfile").GetString());
        Assert.Equal(7, manifestRoot.GetProperty("expectations").GetArrayLength());
        Assert.Equal(3, manifestRoot.GetProperty("artifacts").GetArrayLength());
        Assert.Equal("roundtrip-html", manifestRoot.GetProperty("artifacts")[2].GetProperty("kind").GetString());
        Assert.True(manifestRoot.GetProperty("roundTripScore").GetProperty("score").GetDouble() >= 0D);
        Assert.Contains("HtmlCommentSkipped", manifestJson, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("WordOpenXmlPackageValid", manifestJson, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlArtifactGallery_DefaultImportBlocksLocalFileResources() {
        string artifactDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.HtmlArtifactGallery.Security." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(artifactDirectory);
        string localImagePath = Path.Combine(artifactDirectory, "local-secret.png");
        File.WriteAllBytes(localImagePath, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII="));
        string html = "<html><body><h1>Gallery</h1><img src=\"" + new Uri(localImagePath).AbsoluteUri + "\" alt=\"Local secret\"></body></html>";

        try {
            HtmlCapabilityGalleryManifest manifest = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
                Profile = HtmlConversionProfile.Document,
                Trust = HtmlInputTrust.Trusted
            }).SaveHtmlCapabilityGallery(artifactDirectory, new WordHtmlCapabilityGalleryOptions {
                ScenarioId = "offline-default",
                Title = "Offline Default"
            });

            Assert.Contains(manifest.Result.Diagnostics.Diagnostics,
                diagnostic => diagnostic.Code == "ImageSkippedByPolicy");
            Assert.Contains(manifest.ResourceManifest.Resources,
                resource => resource.Source == new Uri(localImagePath).AbsoluteUri && !resource.IsAllowed);
        } finally {
            Directory.Delete(artifactDirectory, recursive: true);
        }
    }
}
