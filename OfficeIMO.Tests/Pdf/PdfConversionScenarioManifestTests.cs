using System.Text.Json;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionScenarioManifestTests {
    private static readonly string[] RequiredPaths = {
        "word-to-pdf",
        "excel-to-pdf",
        "markdown-to-pdf",
        "html-to-pdf",
        "powerpoint-to-pdf",
        "pdf-to-logical",
        "pdf-to-html"
    };

    [Fact]
    public void PdfConversionManifest_CoversEverySupportedConversionPathWithObservableProof() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(GetManifestPath()));
        JsonElement root = document.RootElement;
        Assert.Equal(1, root.GetProperty("version").GetInt32());

        JsonElement scenarios = root.GetProperty("scenarios");
        Assert.True(scenarios.GetArrayLength() >= RequiredPaths.Length);

        var seenPaths = new HashSet<string>(StringComparer.Ordinal);
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement scenario in scenarios.EnumerateArray()) {
            string id = RequireString(scenario, "id");
            Assert.True(seenIds.Add(id), "Scenario ids must be unique. Duplicate: " + id);

            string path = RequireString(scenario, "path");
            seenPaths.Add(path);
            Assert.Equal("supported", RequireString(scenario, "status"));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "converter")));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "sourceFormat")));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "targetFormat")));
            Assert.NotEmpty(ReadStringArray(scenario, "sourceFeatures"));
            Assert.True(ReadStringArray(scenario, "visualReviewFiles").Count > 0, "Scenario " + id + " needs at least one review artifact.");

            JsonElement proof = scenario.GetProperty("proof");
            Assert.True(proof.GetProperty("hash").GetBoolean(), "Scenario " + id + " must include hash proof.");
            Assert.NotEmpty(ReadStringArray(proof, "textMarkers"));
            Assert.NotEmpty(ReadStringArray(proof, "logicalSignals"));
            Assert.True(proof.GetProperty("visualPages").GetInt32() > 0, "Scenario " + id + " must declare visual page evidence.");
            Assert.NotEmpty(ReadStringArray(proof, "validatorEvidence"));
        }

        foreach (string requiredPath in RequiredPaths) {
            Assert.Contains(requiredPath, seenPaths);
        }
    }

    [Fact]
    public void HtmlDocumentProfile_ProducesManifestedReviewProof() {
        const string linkUri = "https://example.com/pdf-conversion-manifest";
        byte[] pdf = CreatePracticalHtmlSample(linkUri).SaveAsPdf(HtmlPdfSaveOptions.CreateDocumentProfile());
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 2);
        Assert.Contains(logical.TextBlocks, block => block.Text.IndexOf("Practical HTML", StringComparison.Ordinal) >= 0);
        Assert.Contains(logical.TextBlocks, block => block.Text.IndexOf("Second page marker", StringComparison.Ordinal) >= 0);
        Assert.Contains(logical.GetLinksByUri(linkUri), link => link.Contents == "Report link");
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");

        WriteReviewArtifact("practical-html.pdf", pdf);
    }

    [Fact]
    public void PdfLogicalAndHtmlProfiles_ProduceManifestedReadbackProof() {
        byte[] pdf = CreateLogicalProofPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfHtmlConverter.ToHtml(logical, options);

        Assert.Equal(1, logical.PageCount);
        Assert.Contains(logical.Headings, heading => heading.Text == "Logical Heading");
        Assert.Contains(logical.ListItems, item => item.Text == "Detected logical bullet");
        Assert.NotEmpty(logical.Tables);
        Assert.Contains(logical.GetLinksByUri("https://example.com/logical-proof"), link => link.Contents == "Logical PDF sample");
        Assert.Contains(logical.Images, image => image.Width > 0D && image.Height > 0D);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);

        WriteReviewArtifact("pdf-to-html-logical-source.pdf", pdf);
        WriteReviewArtifact("pdf-to-html-positioned-review.html", System.Text.Encoding.UTF8.GetBytes(html));
    }

    [Fact]
    public void TypographyProfile_ProducesMultilingualBusinessReportProof() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf;
        try {
            pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                    CompressContentStreams = false,
                    CompressEmbeddedFonts = false,
                    PageWidth = 520,
                    PageHeight = 420,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36
                })
                .UseFontFamily("OfficeIMO Multilingual", fontPath)
                .Header(header => header.Text("Q2 multilingual revenue report"))
                .H1("Q2 multilingual revenue report")
                .Paragraph(paragraph => paragraph.Text("Executive summary: Zażółć gęślą jaźń Łódź."))
                .Paragraph(paragraph => paragraph.Text("Regional notes: Ελλάδα Athens pipeline; Київ renewal forecast."))
                .Table(new[] {
                    new[] { "Region", "Signal", "Status" },
                    new[] { "Polska", "Łódź", "Ready" },
                    new[] { "Ελλάδα", "Athens", "Ready" },
                    new[] { "Україна", "Київ", "Watch" }
                }, style: new PdfCore.PdfTableStyle {
                    HeaderRowCount = 1,
                    CellPaddingX = 6,
                    CellPaddingY = 4
                })
                .Footer(footer => footer.Text("Generated proof {page}/{pages}"))
                .ToBytes();
        } catch (ArgumentException exception) when (exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal)) {
            return;
        }

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Q2 multilingual revenue report", text, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.Contains("Ελλάδα", text, StringComparison.Ordinal);
        Assert.Contains("Київ", text, StringComparison.Ordinal);

        WriteReviewArtifact("multilingual-business-report.pdf", pdf);
    }

    private static string RequireString(JsonElement element, string propertyName) {
        string? value = element.GetProperty(propertyName).GetString();
        Assert.False(string.IsNullOrWhiteSpace(value), propertyName + " cannot be empty.");
        return value!;
    }

    private static IReadOnlyList<string> ReadStringArray(JsonElement element, string propertyName) {
        var values = new List<string>();
        foreach (JsonElement item in element.GetProperty(propertyName).EnumerateArray()) {
            string? value = item.GetString();
            if (!string.IsNullOrWhiteSpace(value)) {
                values.Add(value!);
            }
        }

        return values;
    }

    private static byte[] CreateLogicalProofPdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical PDF sample", author: "OfficeIMO")
            .H1("Logical Heading", linkUri: "https://example.com/logical-proof", linkContents: "Logical PDF sample")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Logical proof pixel")
            .ToBytes();
    }

    private static string CreatePracticalHtmlSample(string linkUri) {
        string pixel = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        return $$"""
<html>
<head>
  <style>
    table { border-collapse: collapse; }
    td, th { border: 1px solid #444; padding: 4px; }
    .page-two { break-before: page; }
  </style>
</head>
<body>
  <h1>Practical HTML</h1>
  <p><a href="{{linkUri}}">Report link</a></p>
  <p><img src="data:image/png;base64,{{pixel}}" alt="Embedded pixel" width="24" height="24"></p>
  <table>
    <tr><th>Area</th><th>Status</th></tr>
    <tr><td>Table marker</td><td>Ready</td></tr>
  </table>
  <section class="page-two"><h2>Second page marker</h2><p>Page break proof.</p></section>
</body>
</html>
""";
    }

    private static void WriteReviewArtifact(string fileName, byte[] bytes) {
        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, fileName), bytes);
    }

    private static string GetManifestPath() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            string candidate = Path.Combine(directory.FullName, "Docs", "pdf-conversion-scenarios.json");
            if (File.Exists(candidate)) {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException("Could not locate Docs/pdf-conversion-scenarios.json from test runtime base directory.");
    }

}
