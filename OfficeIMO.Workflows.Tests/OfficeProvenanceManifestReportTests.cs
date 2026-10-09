using System.Text.Json;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceManifestReportTests {
    [Fact]
    public void EmbeddedImageCredentialRetainsItsSummaryAndPackageLocation() {
        byte[] png = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "unsigned-credential-12-actions.png"));
        using var output = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var part in new[] {
                ("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"),
                ("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
                ("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>") }) {
                using var writer = new StreamWriter(zip.CreateEntry(part.Item1).Open());
                writer.Write(part.Item2);
            }
            using var image = zip.CreateEntry("word/media/image1.png").Open();
            image.Write(png);
        }
        OfficeProvenanceReport report = OfficeProvenanceBufferWorkflow.Inspect(output.ToArray(), "embedded.docx");
        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence, item => item.Manifest != null);
        Assert.Contains("word/media/image1.png/", evidence.Location);
        Assert.Equal(12, evidence.Manifest!.Actions.Count);
        Assert.Equal("Editor 11", evidence.Manifest.Actions[11].SoftwareAgent);
    }

    [Fact]
    public void CanonicalReportsRetainCredentialActionsBeyondTheCompactPreview() {
        // Synthetic unsigned C2PA fixture: 12 referenced actions and certificate-name fields,
        // with deliberately unverifiable hashes and signature bytes.
        byte[] input = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "unsigned-credential-12-actions.png"));
        OfficeProvenanceReport report = OfficeProvenanceBufferWorkflow.Inspect(input, "image.png");
        OfficeProvenanceWorkflowResult result = OfficeProvenanceReportSerializer.FromBuffer("image.png", input, report);
        using var json = JsonDocument.Parse(OfficeProvenanceReportSerializer.Serialize(result));
        JsonElement summary = json.RootElement.GetProperty("inspection").GetProperty("evidence")[0].GetProperty("manifest");
        Assert.Equal(12, summary.GetProperty("actions").GetArrayLength());
        Assert.Equal("Editor 11", summary.GetProperty("actions")[11].GetProperty("softwareAgent").GetString());
        Assert.Equal("TrainedAlgorithmicMedia", summary.GetProperty("actions")[0].GetProperty("digitalSourceKind").GetString());
        Assert.Equal("2026-10-09T10:00:00Z", summary.GetProperty("actions")[0].GetProperty("when").GetString());
        Assert.Equal("source.png", summary.GetProperty("ingredients")[0].GetString());
        Assert.Equal("Unverified subject", summary.GetProperty("signedBy").GetString());
        Assert.Equal("Unverified issuer", summary.GetProperty("certificateIssuer").GetString());
        Assert.Equal(1, summary.GetProperty("manifestCount").GetInt32());
        Assert.Equal("NotRequested", json.RootElement.GetProperty("checks").GetProperty("verification").GetString());
        var roundTrip = JsonSerializer.Deserialize<ProvenanceResultDto>(json.RootElement.GetRawText(), new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        Assert.Equal("urn:uuid:active", roundTrip!.Inspection!.Evidence[0].Manifest!.Label);
        using var batch = JsonDocument.Parse(OfficeProvenanceReportSerializer.SerializeBatch([result]));
        Assert.Equal(12, batch.RootElement.GetProperty("results")[0].GetProperty("inspection").GetProperty("evidence")[0].GetProperty("manifest").GetProperty("actions").GetArrayLength());
    }
}
