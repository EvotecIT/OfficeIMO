using System.Text.Json;
using System.Text.Json.Nodes;
using Json.Schema;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentAccessibilityTests {
    // Official SARIF 2.1.0 schema (OASIS), as published at https://json.schemastore.org/sarif-2.1.0.json.
    private static readonly Lazy<JsonSchema> SarifSchema = new(() =>
        JsonSchema.FromFile(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Sarif", "sarif-schema-2.1.0.json")));

    [Theory]
    [InlineData("report.docx", "Word")]
    [InlineData("report.xlsx", "Excel")]
    [InlineData("report.pptx", "PowerPoint")]
    public void EveryFormatReturnsTheSharedReportAsSchemaValidSarif(string name, string format) {
        string root = Directory.CreateTempSubdirectory("officeimo-accessibility-").FullName;
        try {
            string input = Path.Combine(root, name);
            Create(input);

            OfficeAccessibilityReport report = OfficeDocument.InspectAccessibility(input);
            Assert.Equal(format, report.Format);
            Assert.Contains(report.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel);

            string sarif = report.ToSarif(input);
            EvaluationResults validation = SarifSchema.Value.Evaluate(JsonNode.Parse(sarif), new EvaluationOptions { OutputFormat = OutputFormat.List });
            Assert.True(validation.IsValid, string.Join("; ", validation.Details.Where(detail => detail.HasErrors)
                .SelectMany(detail => detail.Errors!.Select(error => detail.InstanceLocation + ": " + error.Value))));

            using JsonDocument log = JsonDocument.Parse(sarif);
            JsonElement run = log.RootElement.GetProperty("runs")[0];
            JsonElement rules = run.GetProperty("tool").GetProperty("driver").GetProperty("rules");
            JsonElement results = run.GetProperty("results");
            Assert.Equal(report.Findings.Count, results.GetArrayLength());
            Assert.Equal(new Uri(input).AbsoluteUri, run.GetProperty("artifacts")[0].GetProperty("location").GetProperty("uri").GetString());
            for (int index = 0; index < report.Findings.Count; index++) {
                OfficeAccessibilityFinding finding = report.Findings[index];
                JsonElement result = results[index];
                Assert.Equal(finding.Code, result.GetProperty("ruleId").GetString());
                Assert.Equal(finding.Code, rules[result.GetProperty("ruleIndex").GetInt32()].GetProperty("id").GetString());
                Assert.Equal(finding.Severity switch {
                    OfficeAccessibilitySeverity.Error => "error",
                    OfficeAccessibilitySeverity.Warning => "warning",
                    _ => "note"
                }, result.GetProperty("level").GetString());
                Assert.Equal(finding.Location, result.GetProperty("properties").GetProperty("location").GetString());
            }
            Assert.Equal(report.ErrorCount, run.GetProperty("properties").GetProperty("errorCount").GetInt32());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void SarifWithoutAnArtifactUsesLogicalLocationsAndStaysValid() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Q1 Sales");
        sheet.Cell(1, 1).SetValue("Total").SetColor(OfficeColor.FromRgb(0xDD, 0xDD, 0xDD));
        sheet.Range("A1:B1").Merge();
        OfficeAccessibilityReport report = OfficeDocument.InspectAccessibility(workbook, OfficeAccessibilityPolicyProfile.Strict);

        string sarif = report.ToSarif(indented: false);

        Assert.True(SarifSchema.Value.Evaluate(JsonNode.Parse(sarif)).IsValid);
        Assert.False(SarifSchema.Value.Evaluate(JsonNode.Parse(sarif.Replace("\"level\":\"error\"", "\"level\":\"fatal\""))).IsValid);
        using JsonDocument log = JsonDocument.Parse(sarif);
        JsonElement run = log.RootElement.GetProperty("runs")[0];
        Assert.False(run.TryGetProperty("artifacts", out _));
        JsonElement contrast = run.GetProperty("results").EnumerateArray()
            .Single(result => result.GetProperty("ruleId").GetString() == OfficeDiagnosticCodes.AccessibilityLowContrast);
        JsonElement location = contrast.GetProperty("locations")[0];
        Assert.False(location.TryGetProperty("physicalLocation", out _));
        Assert.Equal("'Q1 Sales'!A1", location.GetProperty("logicalLocations")[0].GetProperty("fullyQualifiedName").GetString());
        Assert.Equal(4.5D, contrast.GetProperty("properties").GetProperty("requiredValue").GetDouble());
        Assert.Equal("Strict", run.GetProperty("properties").GetProperty("profile").GetString());
        JsonElement rule = run.GetProperty("tool").GetProperty("driver").GetProperty("rules").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == OfficeDiagnosticCodes.AccessibilityMergedCells);
        Assert.Equal("MergedCells", rule.GetProperty("name").GetString());
        Assert.Equal("error", rule.GetProperty("defaultConfiguration").GetProperty("level").GetString());
    }

    [Fact]
    public void RelativeArtifactPathsStayRelative() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddSheet("Data").Cell(1, 1).SetValue("x");
        workbook.Sheets[0].Range("A1:B1").Merge();
        using JsonDocument log = JsonDocument.Parse(OfficeDocument.InspectAccessibility(workbook).ToSarif(Path.Combine("docs", "Q1 report.xlsx")));
        Assert.Equal("docs/Q1%20report.xlsx", log.RootElement.GetProperty("runs")[0].GetProperty("artifacts")[0]
            .GetProperty("location").GetProperty("uri").GetString());
    }

    [Fact]
    public void UnsupportedInputsAreRejected() {
        Assert.False(OfficeDocument.SupportsAccessibility("notes.pdf"));
        Assert.True(OfficeDocument.SupportsAccessibility("Budget.XLSX"));
        string root = Directory.CreateTempSubdirectory("officeimo-accessibility-").FullName;
        try {
            string text = Path.Combine(root, "notes.txt");
            File.WriteAllText(text, "plain");
            Assert.Throws<NotSupportedException>(() => OfficeDocument.InspectAccessibility(text));
            Assert.Throws<NotSupportedException>(() => OfficeDocument.InspectAccessibility(new object()));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ValidationRunsTheExcelCheckerWithItsCodes() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddSheet("Sheet1").Cell(1, 1).SetValue("Heading");
        workbook.Sheets[0].Range("A1:B1").Merge();

        OfficeValidationReport report = OfficeDocument.Validate(workbook, new[] { OfficeDocumentCheck.Accessibility },
            new OfficeDocumentValidationOptions { ExcelAccessibility = new ExcelAccessibilityOptions { CheckDefaultSheetNames = false } });

        Assert.Contains(report.NativeReports, evidence => evidence is OfficeAccessibilityReport { Format: "Excel" });
        OfficeDiagnostic merged = Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.AccessibilityMergedCells);
        Assert.Equal("Sheet1!A1:B1", merged.Location);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.AccessibilityDefaultSheetName);
    }

    private static void Create(string path) {
        switch (Path.GetExtension(path)) {
            case ".docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph("Report");
                    document.AddHyperLink("click here", new Uri("https://officeimo.com"));
                    document.Save();
                }
                break;
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = document.AddSheet("Sheet1");
                    sheet.Cell(1, 1).SetValue("Total");
                    sheet.Range("A1:B1").Merge();
                    sheet.SetHyperlink("A3", "https://officeimo.com", "here");
                    document.Save();
                }
                break;
            default:
                using (PowerPointDocument document = PowerPointDocument.Create(path)) {
                    PowerPointSlide slide = document.AddSlide();
                    PowerPointTextBox link = slide.AddTextBox("click here", OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(180), OfficeLength.Points(28));
                    link.Paragraphs.Single().Runs.Single().SetHyperlink("https://officeimo.com");
                    document.Save();
                }
                break;
        }
    }
}
