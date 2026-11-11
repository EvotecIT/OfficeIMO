using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class HiddenDataCommandTests {
    [Fact]
    public async Task RemoveWritesACleanCopyAndKeepsProvenanceUnlessItIsSelected() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tool-hidden-data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "shared.docx");
            string image = Path.Combine(directory, "generated.jpg");
            byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(jpeg);
            metadata.XmpProfile = Encoding.UTF8.GetBytes(
                    "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                    "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
                    "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
                    "</rdf:Description></rdf:RDF></x:xmpmeta>");
            File.WriteAllBytes(image, OfficeImageMetadata.Apply(jpeg, metadata));
            using (WordDocument document = WordDocument.Create(input)) {
                document.Properties.Creator = "Private Author";
                document.AddParagraph().AddImage(image);
                document.Save();
            }
            byte[] original = File.ReadAllBytes(input);

            (int exitCode, string output) kept = await RunAsync("hidden-data", "remove", input);
            Assert.Equal(0, kept.exitCode);
            using (JsonDocument json = JsonDocument.Parse(kept.output)) {
                JsonElement[] findings = json.RootElement.GetProperty("findings").EnumerateArray().ToArray();
                Assert.Contains(findings, finding => finding.GetProperty("kind").GetString() == "DocumentProperties" && finding.GetProperty("outcome").GetString() == "Removed");
                JsonElement provenance = Assert.Single(findings, finding => finding.GetProperty("kind").GetString() == "Provenance");
                Assert.Equal("Retained", provenance.GetProperty("outcome").GetString());
                Assert.Equal("TrainedAlgorithmicMedia", provenance.GetProperty("digitalSourceKind").GetString());
            }
            string cleaned = Path.Combine(directory, "shared.cleaned.docx");
            Assert.True(OfficeProvenanceInspector.InspectFile(cleaned).HasGenerativeAiDeclaration);
            Assert.Equal(original, File.ReadAllBytes(input));

            Assert.Equal(6, (await RunAsync("hidden-data", "remove", input)).ExitCode);
            Assert.Equal(0, (await RunAsync("hidden-data", "remove", input, "--force", "--kinds", "default,provenance")).ExitCode);
            Assert.Empty(OfficeProvenanceInspector.InspectFile(cleaned).Evidence);
            Assert.Equal(2, (await RunAsync("hidden-data", "remove", input, "--kinds", "everything")).ExitCode);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static async Task<(int ExitCode, string Output)> RunAsync(params string[] args) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()) + error);
    }
}
