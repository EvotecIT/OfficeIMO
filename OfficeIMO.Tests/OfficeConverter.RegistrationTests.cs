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
    public async Task RegisteredRtfBatchResumesAndRejectsChangedPhysicalSettingsAndAcceptancePolicy() {
        OfficeWorkflow.RegisterConversions();
        string root = Directory.CreateTempSubdirectory("officeimo-registered-batch-test-").FullName;
        try {
            string source = Path.Combine(root, "input.rtf"), output = Path.Combine(root, "output");
            File.WriteAllText(source, @"{\rtf1\ansi Checkpointed RTF route}", Encoding.ASCII);
            var options = new OfficeBatchOptions { InputPaths = new[] { source }, OutputDirectory = output,
                ResumeFrom = Path.Combine(output, ".batch.json"), TargetExtension = ".pdf",
                ConversionOptions = new OfficeConversionOptions { RegisteredSettings = new OfficeWorkflowRouteSettings {
                    Rtf = new RtfToPdfOptions { PdfOptions = new PdfCreateOptions { DefaultFontSize = OfficeLength.Points(14), Theme = OfficeTheme.Modern,
                        DefaultTextColor = OfficeColor.FromNormalizedRgb(0.123456, 0.3, 0.6) } }
                } } };
            OfficeBatchReport first = await OfficeConverter.ConvertAsync(options);
            Assert.True(first.Summary.Completed == 1, string.Join("\n", first.Items.Select(item => item.Summary)));
            OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(options);
            Assert.Equal(1, resumed.Summary.Reused);
            byte[] original = File.ReadAllBytes(Path.Combine(output, "input.pdf"));
            OfficeConversionOptions changedSize = options.ConversionOptions.Clone();
            var changedRoutes = Assert.IsType<OfficeWorkflowRouteSettings>(changedSize.RegisteredSettings);
            changedRoutes.Rtf!.PdfOptions!.DefaultFontSize = OfficeLength.Emus(OfficeLength.Points(14).ToEmus() + 1);
            OfficeBatchReport changed = await OfficeConverter.ConvertAsync(options with { ConversionOptions = changedSize });
            Assert.Equal(1, changed.Summary.Failed);
            Assert.Equal(0, changed.Summary.Reused);
            Assert.Equal(original, File.ReadAllBytes(Path.Combine(output, "input.pdf")));
            OfficeConversionOptions changedColor = options.ConversionOptions.Clone();
            var colorRoutes = Assert.IsType<OfficeWorkflowRouteSettings>(changedColor.RegisteredSettings);
            colorRoutes.Rtf!.PdfOptions!.DefaultTextColor = OfficeColor.FromNormalizedRgb(0.12345601, 0.3, 0.6);
            OfficeBatchReport colorRejected = await OfficeConverter.ConvertAsync(options with { ConversionOptions = changedColor });
            Assert.Equal(1, colorRejected.Summary.Failed);
            Assert.Equal(0, colorRejected.Summary.Reused);
            Assert.Equal(original, File.ReadAllBytes(Path.Combine(output, "input.pdf")));
            OfficeConversionOptions changedPolicy = options.ConversionOptions.Clone();
            changedPolicy.FailOnLoss = true;
            OfficeBatchReport rejected = await OfficeConverter.ConvertAsync(options with { ConversionOptions = changedPolicy });
            Assert.Equal(1, rejected.Summary.Failed);
            Assert.Equal(0, rejected.Summary.Reused);
        } finally { Directory.Delete(root, true); }
    }
}
