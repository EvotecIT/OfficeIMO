using System.Buffers.Binary;
using System.Net;
using System.Text.Json;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed class HtmlCredentialSummaryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task FileWorkflowRetainsDirectAndNestedCredentialSummaries(bool image, bool srcdoc) {
        byte[] png = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "unsigned-credential-12-actions.png"));
        string html = image
            ? "<html><body><img src=\"data:image/png;base64," + Convert.ToBase64String(png) + "\"></body></html>"
            : "<html><head><script type=\"application/c2pa\">" + Convert.ToBase64String(ManifestPayload(png)) + "</script></head><body></body></html>";
        if (srcdoc) html = "<html><body><iframe srcdoc=\"" + WebUtility.HtmlEncode(html) + "\"></iframe></body></html>";
        string path = Path.Combine(Path.GetTempPath(), "officeimo-html-credential-" + Guid.NewGuid().ToString("N") + ".html");
        try {
            File.WriteAllText(path, html);
            OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
                new OfficeProvenanceWorkflowRequest { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Inspect });
            OfficeProvenanceEvidence evidence = Assert.Single(result.Inspection!.Evidence, item => item.Carrier == OfficeProvenanceCarrierKind.C2paManifest);
            OfficeC2paManifestSummary summary = Assert.IsType<OfficeC2paManifestSummary>(evidence.Manifest);
            Assert.Equal("urn:uuid:active", summary.Label);
            Assert.Equal(12, summary.Actions.Count);
            Assert.Equal("Editor 11", summary.Actions[11].SoftwareAgent);
            Assert.Equal("Unverified subject", summary.SignedBy);
            Assert.Equal("Unverified issuer", summary.CertificateIssuer);
            if (srcdoc) Assert.Contains("iframe[srcdoc]", evidence.Location);
            if (image) Assert.Contains("img[src]", evidence.Location);
            using var json = JsonDocument.Parse(OfficeProvenanceReportSerializer.Serialize(result));
            JsonElement manifest = json.RootElement.GetProperty("inspection").GetProperty("evidence")
                .EnumerateArray().Single(item => item.GetProperty("manifest").ValueKind == JsonValueKind.Object).GetProperty("manifest");
            Assert.Equal(12, manifest.GetProperty("actions").GetArrayLength());
            Assert.Equal("Editor 11", manifest.GetProperty("actions")[11].GetProperty("softwareAgent").GetString());
            Assert.Equal("NotRequested", json.RootElement.GetProperty("checks").GetProperty("verification").GetString());
        } finally {
            File.Delete(path);
        }
    }

    private static byte[] ManifestPayload(byte[] png) {
        for (int offset = 8; offset + 12 <= png.Length;) {
            int length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset, 4));
            if (png.AsSpan(offset + 4, 4).SequenceEqual("caBX"u8)) return png.AsSpan(offset + 8, length).ToArray();
            offset += length + 12;
        }
        throw new InvalidDataException("The credential fixture has no C2PA carrier.");
    }
}
