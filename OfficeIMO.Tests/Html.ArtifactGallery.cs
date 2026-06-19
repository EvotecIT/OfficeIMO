using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml() {
        const string html = """
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
                    <!-- skipped comments are diagnostic evidence, not visible document content -->
                </article>
            </body>
            </html>
            """;

        string artifactDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.HtmlArtifactGallery");
        Directory.CreateDirectory(artifactDirectory);
        string inputPath = Path.Combine(artifactDirectory, "quarterly-report.input.html");
        string docxPath = Path.Combine(artifactDirectory, "quarterly-report.docx");
        string roundTripPath = Path.Combine(artifactDirectory, "quarterly-report.roundtrip.html");
        string manifestPath = Path.Combine(artifactDirectory, "quarterly-report.manifest.md");

        var importOptions = HtmlToWordOptions.CreateTrustedDocumentProfile();
        importOptions.EnableAccessibilityDiagnostics = true;
        importOptions.ConversionReport.Clear();
        using var document = html.LoadFromHtml(importOptions);
        using MemoryStream packageStream = document.SaveAsMemoryStream();

        File.WriteAllText(inputPath, html);
        File.WriteAllBytes(docxPath, packageStream.ToArray());
        packageStream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(packageStream, false);
        var errors = new OpenXmlValidator().Validate(package).ToList();
        Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));

        string roundTripHtml = document.ToHtml(new WordToHtmlOptions {
            IncludeListStyles = true,
            IncludeTableColumnGroups = true,
            IncludeDefaultCss = true
        });
        File.WriteAllText(roundTripPath, roundTripHtml);

        var scenario = new HtmlCapabilityGalleryScenario(
            "quarterly-report",
            "Quarterly Report",
            "Word HTML",
            "Validates HTML import, DOCX package validity, round-trip HTML export, form controls, tables, and diagnostics.");
        var result = new HtmlCapabilityGalleryResult(scenario);
        result.AddArtifact(HtmlCapabilityGalleryArtifact.FromFile("source", "input-html", inputPath, "text/html"));
        result.AddArtifact(HtmlCapabilityGalleryArtifact.FromFile("docx", "docx", docxPath, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        result.AddArtifact(HtmlCapabilityGalleryArtifact.FromFile("roundtrip", "roundtrip-html", roundTripPath, "text/html"));
        result.Diagnostics.AddRange(importOptions.ConversionReport.Diagnostics);
        File.WriteAllText(manifestPath, BuildManifest(result));

        Assert.Contains("<h1>Quarterly Report</h1>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<thead>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<tfoot>", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"checkbox\"", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<select", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("skipped comments are diagnostic evidence", roundTripHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(importOptions.Diagnostics, diagnostic => string.Equals(diagnostic.Code, "HtmlCommentSkipped", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(importOptions.ConversionReport.Diagnostics, diagnostic => string.Equals(diagnostic.Code, "HtmlCommentSkipped", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("quarterly-report", result.Scenario.Id);
        Assert.Equal(3, result.Artifacts.Count);
        Assert.Contains(result.Artifacts, artifact => artifact.Kind == "docx" && artifact.Length > 0 && artifact.Sha256.Length == 64);
        Assert.Contains(result.Diagnostics.Diagnostics, diagnostic => diagnostic.Component == "OfficeIMO.Word.Html" && diagnostic.Code == "HtmlCommentSkipped");
        Assert.True(File.Exists(inputPath));
        Assert.True(File.Exists(docxPath));
        Assert.True(File.Exists(roundTripPath));
        Assert.True(File.Exists(manifestPath));
        Assert.Contains("HtmlCommentSkipped", File.ReadAllText(manifestPath), StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildManifest(HtmlCapabilityGalleryResult result) {
        var builder = new StringBuilder();
        builder.AppendLine("# HTML Capability Gallery Scenario");
        builder.AppendLine();
        builder.AppendLine($"Id: {result.Scenario.Id}");
        builder.AppendLine($"Title: {result.Scenario.Title}");
        builder.AppendLine($"Category: {result.Scenario.Category}");
        builder.AppendLine($"Description: {result.Scenario.Description}");
        builder.AppendLine();
        builder.AppendLine("## Artifacts");
        foreach (HtmlCapabilityGalleryArtifact artifact in result.Artifacts) {
            builder.AppendLine($"- {artifact.Kind}: {Path.GetFileName(artifact.Path)} ({artifact.MediaType}, {artifact.Length} bytes, sha256={artifact.Sha256})");
        }

        builder.AppendLine();
        builder.AppendLine("## Diagnostics");
        foreach (HtmlDiagnostic diagnostic in result.Diagnostics.Diagnostics) {
            builder.AppendLine($"- {diagnostic.Component}:{diagnostic.Code}:{diagnostic.Severity}: {diagnostic.Message}");
        }

        return builder.ToString();
    }
}
