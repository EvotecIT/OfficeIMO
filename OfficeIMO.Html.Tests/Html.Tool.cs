#if NET8_0_OR_GREATER
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Tool;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfToolTests {
    [Fact]
    public async Task HtmlTool_ConvertsStandardInputToPdfStandardOutput() {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes(
            "<html lang='en'><head><title>Tool</title></head><body><h1>Command line</h1><p>Ready</p></body></html>"));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await HtmlPdfToolApp.RunAsync(
            new[] { "convert", "-", "--input-format", "html", "--output", "-" },
            input,
            output,
            error);

        Assert.True(exitCode == 0, error.ToString());
        Assert.True(output.Length > 4);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(output.ToArray(), 0, 4));
        Assert.Contains("Command line", OfficeIMO.Pdf.PdfReadDocument.Open(output.ToArray()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlTool_ConvertsMhtmlStandardInputWithEmbeddedResources() {
        var archive = new MhtmlDocument(
            "<html lang='en'><head><title>Archive</title><link rel='stylesheet' href='cid:style'></head><body><h1>MHTML command line</h1><p>resource</p></body></html>",
            new[] {
                new MhtmlResource(
                    Encoding.UTF8.GetBytes("p::before{content:'Embedded '}"),
                    "text/css",
                    contentId: "style",
                    fileName: "style.css")
            });
        await using var input = new MemoryStream(archive.ToBytes());
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await HtmlPdfToolApp.RunAsync(
            new[] { "convert", "-", "--input-format", "mhtml", "--output", "-" },
            input,
            output,
            error);

        Assert.True(exitCode == 0, error.ToString());
        byte[] pdf = output.ToArray();
        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
        Assert.Contains("MHTML command line", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Embedded resource", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("StylesheetResourceUnavailable", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlTool_EmitsMachineReadableCapabilityContract() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await HtmlPdfToolApp.RunAsync(
            new[] { "capabilities", "--format", "json" },
            input,
            output,
            error);

        Assert.Equal(0, exitCode);
        using JsonDocument json = JsonDocument.Parse(output.ToArray());
        Assert.Equal(JsonValueKind.Array, json.RootElement.ValueKind);
        Assert.Contains(json.RootElement.EnumerateArray(), item =>
            item.GetProperty("id").GetString() == "css-length-math"
            && item.GetProperty("supportLevel").GetString() == "Full");
    }

    [Fact]
    public async Task HtmlTool_AcceptsTheExplicitTextCapabilityFormat() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await HtmlPdfToolApp.RunAsync(
            new[] { "capabilities", "--format", "text" },
            input,
            output,
            error);

        Assert.Equal(0, exitCode);
        Assert.Contains("css-length-math", Encoding.UTF8.GetString(output.ToArray()), StringComparison.Ordinal);
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public async Task HtmlTool_PdfUaStatusIsBoundToTheExactBytesWrittenToStandardOutput() {
        string fontPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".ttf");
        await File.WriteAllBytesAsync(
            fontPath,
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(
                "Accessible tool outputProofExact artifact".Distinct().Select(static character => (int)character).ToArray()));
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes(
            "<html lang='fr'><head><title>Accessible tool output</title></head><body><h1>Proof</h1><p>Exact artifact</p></body></html>"));
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        try {
            int exitCode = await HtmlPdfToolApp.RunAsync(
                new[] {
                    "convert", "-", "--input-format", "html", "--output", "-",
                    "--pdf-ua-language", "en-US",
                    "--font-family", "Tool Contract",
                    "--font-regular", fontPath
                },
                input,
                output,
                error);

            Assert.Equal(0, exitCode);
            byte[] artifact = output.ToArray();
            Assert.Equal("en-US", OfficeIMO.Pdf.PdfReadDocument.Open(artifact).CatalogLanguage);
            OfficeIMO.Pdf.PdfComplianceProofReport proof = OfficeIMO.Pdf.PdfDocument.Open(artifact)
                .AssessComplianceProof(OfficeIMO.Pdf.PdfComplianceProfile.PdfUa1);
            Assert.True(proof.HasArtifactEvidence);
            Assert.Equal(artifact.LongLength, proof.ArtifactSizeBytes);
            Assert.Contains("MissingExternalValidation", error.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain("MissingArtifactEvidence", error.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain("InternalGaps", error.ToString(), StringComparison.Ordinal);
        } finally {
            File.Delete(fontPath);
        }
    }

    [Fact]
    public async Task HtmlTool_RequiresAnExplicitFormatForStandardInput() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await HtmlPdfToolApp.RunAsync(
            new[] { "convert", "-", "--output", "-" },
            input,
            output,
            error);

        Assert.Equal(2, exitCode);
        Assert.Contains("--input-format", error.ToString(), StringComparison.Ordinal);
        Assert.Equal(0, output.Length);
    }
}
#endif
