using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPostprocessingFailureEvidenceTests {
    [Theory]
    [InlineData(".odg", "odg-pdf")]
    [InlineData(".vdx", "visio-pdf")]
    [InlineData(".txt", "txt-pdf")]
    [InlineData(".docx", "docx-pdf")]
    public async Task DeferredEncryptionLimitRetainsConversionReportsAndPreservesDestination(string extension, string route) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-postprocess-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            if (extension == ".odg") {
                var source = OdgDocument.Create();
                source.AddPage("Page").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 2), "Caption");
                source.Save(input);
            } else if (extension == ".vdx") {
                var source = VisioDocument.Create();
                source.AddPage("Page").AddRectangle(1, 1, 2, 1, "Caption");
                source.SaveLegacyXml(input, allowOmissions: true);
            } else if (extension == ".docx") {
                using var source = WordDocument.Create(input);
                source.AddParagraph("Caption");
                source.Save();
            } else {
                File.WriteAllText(input, "Caption");
            }
            byte[] originalSource = File.ReadAllBytes(input);
            var runner = new OfficeWorkflowRunner();
            OfficeWorkflowConversionOptions Options(bool compressed, bool encrypted) {
                var pdf = new PdfOptions();
                if (encrypted) pdf.SetEncryption(new PdfStandardEncryptionOptions("reader") { OwnerPassword = "owner" });
                return new() {
                    CompressPdfOutput = compressed,
                    Draw = extension == ".odg" ? new OdgToPdfOptions { PdfOptions = pdf } : null,
                    Visio = extension == ".vdx" ? new VisioToPdfOptions { Mode = VisioPdfProjectionMode.DiagramPages, PdfOptions = pdf } : null,
                    PlainText = extension == ".txt" ? new PdfPlainTextOptions { PdfOptions = pdf } : null,
                    Word = extension == ".docx" ? new WordToPdfOptions { PdfOptions = pdf } : null
                };
            }
            async Task<OfficeWorkflowResult> Run(bool compressed, bool encrypted, long limit) => await runner.RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output, ConversionRouteId = route,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace, Limits = new() { MaximumOutputBytes = limit },
                ConversionOptions = Options(compressed, encrypted)
            });
            var plain = await Run(false, false, 1024 * 1024);
            Assert.True(plain.Succeeded, plain.Summary);
            long maximum = new FileInfo(output).Length;
            var compressed = await Run(true, false, maximum);
            Assert.True(compressed.Succeeded, compressed.Summary);
            byte[] compressedBytes = File.ReadAllBytes(output);
            byte[] encrypted = PdfSecurityEditor.Encrypt(compressedBytes, new("reader") { OwnerPassword = "owner" }).Pdf;
            Assert.True(encrypted.LongLength > maximum, "The fixture must reach the deferred encryption output limit after successful compression.");
            var accepted = await Run(true, true, 1024 * 1024);
            Assert.True(accepted.Succeeded, accepted.Summary);
            byte[] acceptedBytes = File.ReadAllBytes(output);
            var opened = PdfDocument.Load(acceptedBytes, new PdfLoadOptions { Password = "reader" });
            Assert.True(opened.Inspect().Security.HasEncryption);
            Assert.Contains("Caption", PdfReadDocument.Open(acceptedBytes, new PdfLoadOptions { Password = "reader" }).ExtractText());
            byte[] sentinel = [1, 2, 3]; File.WriteAllBytes(output, sentinel);

            var failed = await Run(true, true, maximum);

            Assert.False(failed.Succeeded);
            Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, failed.FailureKind);
            Assert.Contains(failed.Diagnostics, d => d.Code == "PdfOutputCompression");
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(failed.ConversionEvidence);
            if (extension is ".odg" or ".vdx") Assert.Equal("1", evidence.Facts["sourcePages"]);
            Assert.Equal(
                compressed.ConversionEvidence!.FidelityDiagnostics.Select(d => (d.Source, d.Code, d.Message, d.LossKind, d.Location)),
                evidence.FidelityDiagnostics.Select(d => (d.Source, d.Code, d.Message, d.LossKind, d.Location)));
            foreach (var finding in evidence.FidelityDiagnostics)
                Assert.Single(failed.Diagnostics, d => d.Code == finding.Code && d.Message == finding.Message);
            Assert.Equal(originalSource, File.ReadAllBytes(input));
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }
}
