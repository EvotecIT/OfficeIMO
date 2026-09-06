#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalEngineProofTests {
    [Theory]
    [InlineData("canvas")]
    [InlineData("effect")]
    [InlineData("nested")]
    [InlineData("drawing")]
    public void LogicalTextReplacesPaintExactlyOnceInIndependentExtraction(string scenario) {
        const string replacement = "LogicalReplacement";
        const string paint = "PaintOnly";
        var document = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }).TaggedPdfCatalogMarkers();
        if (scenario == "drawing") {
            var ink = new OfficeDrawing(120D, 40D).AddText(paint, 5D, 5D, 100D, 25D, new OfficeFontInfo("Helvetica", 12D));
            var drawing = new OfficeDrawing(120D, 40D).AddActualTextDrawing(replacement, ink, 5D, 5D);
            document.Canvas(canvas => canvas.Drawing(drawing, 10D, 10D, 120D, 40D));
        } else {
            document.Canvas(canvas => canvas.ActualText(replacement, 10D, 10D, logical => {
                if (scenario == "effect") logical.Effect(OfficeTransform.Identity, .5D, effect => effect.Text(paint, 10D, 10D, 100D, 20D));
                else if (scenario == "nested") logical.ActualText("NestedReplacement", nested => nested
                    .Text(paint, 10D, 10D, 100D, 20D).SearchableText("NestedSearch", 10D, 10D));
                else logical.Text(paint, 10D, 10D, 100D, 20D);
            }));
        }
        byte[] pdf = document.ToBytes();
        Assert.Equal(replacement, PdfReadDocument.Open(pdf).ExtractText().Trim());
        PdfExternalValidator validator = PdfExternalValidator.PopplerText();
        if (!validator.IsAvailable) {
            Assert.NotEqual("1", Environment.GetEnvironmentVariable("OFFICEIMO_REQUIRE_PDF_TEXT_VALIDATOR"));
            return;
        }
        PdfExternalProcessResult result = validator.Run(pdf, "logical-text.pdf");
        Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
        Assert.Equal(replacement, result.Output.Trim());
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        string name = "logical-text-" + scenario;
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), pdf);
        File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(new {
            Scenario = scenario, PdfFile = name + ".pdf", PdfLength = pdf.Length,
            PdfSha256 = Convert.ToHexString(SHA256.HashData(pdf)).ToLowerInvariant(),
            result.ValidatorName, result.ExitCode, ExpectedText = replacement,
            ActualText = result.Output.Trim(), Passed = true
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
#endif
