namespace OfficeIMO.Tests;

public sealed class OfficeConverterRegisteredRouteTests {
    [Fact]
    public async Task RegisteredConversionMustProduceAReadableArtifactBeforePublication() {
        OfficeConverter.Register(new InvalidPdfRoute());
        string directory = Directory.CreateTempSubdirectory("officeimo-registered-test-").FullName;
        try {
            string source = Path.Combine(directory, "source.adoc"), destination = Path.Combine(directory, "result.pdf");
            await File.WriteAllTextAsync(source, "= Source");
            Exception? error = await Record.ExceptionAsync(() => OfficeConverter.ConvertAsync(source, destination));
            Assert.NotNull(error);
            Assert.False(File.Exists(destination));
            Assert.Single(Directory.EnumerateFiles(directory));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private sealed class InvalidPdfRoute : IOfficeConversionRoute {
        public string Id => "asciidoc-pdf";
        public async Task<OfficeConversionReport> ConvertAsync(Stream source, Stream destination, OfficeConversionOptions options, CancellationToken cancellationToken) {
            await destination.WriteAsync(System.Text.Encoding.ASCII.GetBytes("Invalid PDF"), cancellationToken);
            return new OfficeConversionReport("AsciiDoc", "Pdf", itemsWritten: 1);
        }
    }
}
