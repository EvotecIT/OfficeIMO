using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ImportedMobiWorkflowTests {
    [Fact]
    public async Task EncryptedPdfUsesRendererCredentialsForOutputValidation() {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour.azw3");
        using var directory = new TemporaryOutputDirectory();
        string output = Path.Combine(directory.Path, "encrypted.pdf");
        var html = new HtmlPdfExportOptions();
        html.PdfOptions.SetEncryption("reader", "owner");

        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = "ebook-mobi-pdf",
            InputPath = source,
            OutputPath = output,
            ConversionOptions = new OfficeConversionOptions { Html = html }
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Throws<PdfPasswordRequiredException>(() => {
            using PdfDocument locked = PdfDocument.Load(output);
            _ = locked.Inspect();
        });
        var readOptions = new PdfLoadOptions { Password = "reader" };
        using PdfDocument reopened = PdfDocument.Load(output, readOptions);
        Assert.True(reopened.Inspect(readOptions).PageCount > 0);
    }

    [Fact]
    public async Task StrictPdfRenderKeepsImportFindingsAndLeavesTheDestinationIntact() {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour.azw3");
        using var directory = new TemporaryOutputDirectory();
        string output = Path.Combine(directory.Path, "book.pdf");
        await File.WriteAllTextAsync(output, "Existing destination");
        var html = new HtmlPdfExportOptions { FailOnLoss = true };
        html.ResourcePolicy.AllowEmbeddedPackageResources = false;

        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            OfficeConverter.ConvertAsync(source, output, new OfficeConversionOptions {
                RouteId = "ebook-mobi-pdf", Html = html, FileConflictPolicy = OfficePublicationPolicy.Replace
            }));

        Assert.Contains(error.Report.Diagnostics, item => item.Code == "MOBI_FLOWS_OMITTED");
        Assert.Contains(error.Report.Diagnostics, item => item.Format != "OfficeIMO.Mobi" && item.LossKind != OfficeConversionLossKind.None);
        Assert.Equal("Existing destination", await File.ReadAllTextAsync(output));
    }

    private sealed class TemporaryOutputDirectory : IDisposable {
        internal string Path { get; } = Directory.CreateTempSubdirectory("officeimo-mobi-route-").FullName;

        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
