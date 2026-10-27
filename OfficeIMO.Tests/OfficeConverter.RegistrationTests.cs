using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf.Pdf;

using OfficeIMO.Workflows;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterRegistrationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RegisteredRtfRouteConvertsAndReopensEncryptedOutput(bool compress) {
        OfficeWorkflow.RegisterConversions();
        using var input = new MemoryStream(Encoding.ASCII.GetBytes(@"{\rtf1\ansi Registered RTF route}"));
        using var output = new MemoryStream();
        const string password = "synthetic-rtf-password";
        OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, OfficeDocumentFormat.Rtf, output, OfficeDocumentFormat.Pdf,
            new OfficeConversionOptions { CompressPdfOutput = compress, RegisteredSettings = new OfficeWorkflowRouteSettings {
                Rtf = new RtfToPdfOptions { PdfOptions = new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions(password)) }
            } });
        Assert.Equal(1, report.ItemsWritten);
        Assert.Equal(compress, report.Diagnostics.Any(item => item.Code == OfficeDiagnosticCodes.PdfOutputCompressed));
        using var document = PdfDocument.Load(new MemoryStream(output.ToArray()), new PdfLoadOptions { Password = password });
        Assert.Contains("Registered RTF route", document.ExtractText());
    }

    [Fact]
    public async Task RegisteredRtfBatchResumesAndRejectsChangedAcceptancePolicy() {
        OfficeWorkflow.RegisterConversions();
        string root = Directory.CreateTempSubdirectory("officeimo-registered-batch-test-").FullName;
        try {
            string source = Path.Combine(root, "input.rtf"), output = Path.Combine(root, "output");
            File.WriteAllText(source, @"{\rtf1\ansi Checkpointed RTF route}", Encoding.ASCII);
            var options = new OfficeBatchOptions { InputPaths = new[] { source }, OutputDirectory = output,
                ResumeFrom = Path.Combine(output, ".batch.json"), TargetExtension = ".pdf" };
            OfficeBatchReport first = await OfficeConverter.ConvertAsync(options);
            Assert.True(first.Summary.Completed == 1, string.Join("\n", first.Items.Select(item => item.Summary)));
            OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(options);
            Assert.Equal(1, resumed.Summary.Reused);
            OfficeBatchReport rejected = await OfficeConverter.ConvertAsync(options with {
                ConversionOptions = new OfficeConversionOptions { FailOnLoss = true }
            });
            Assert.Equal(1, rejected.Summary.Failed);
            Assert.Equal(0, rejected.Summary.Reused);
        } finally { Directory.Delete(root, true); }
    }
}
