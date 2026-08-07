using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioPdfApiConsistencyTests {
    [Fact]
    public void LoadedDocumentProjectsDirectlyWithoutReopeningOrChangingItsDestination() {
        VisioDocument diagram = CreateDiagram();

        string? destination = diagram.FilePath;
        OfficeDocumentReadResult normalized = diagram.ToOfficeDocumentReadResult();
        PdfDocumentConversionResult conversion = diagram.ToPdfDocumentResult();
        string text = PdfReadDocument.Open(conversion.ToBytes()).ExtractText();

        Assert.Equal(destination, diagram.FilePath);
        Assert.Equal("document.vsdx", normalized.Source.Path);
        Assert.False(string.IsNullOrWhiteSpace(normalized.Source.SourceId));
        Assert.Contains("Gateway service", text, StringComparison.Ordinal);
        Assert.Contains(conversion.Warnings, static warning => warning.Code == "pdf-projection-visio-semantic-fallback");
        Assert.Single(normalized.Pages);
    }

    [Fact]
    public void VisioOwnsTheNeutralModelAndThePdfAdapterHasNoReaderAssemblyDependency() {
        VisioDocument diagram = CreateDiagram();

        OfficeDocumentModel model = diagram.ToOfficeDocumentModel("topology.vsdx");
        OfficeDocumentModelPage page = Assert.Single(model.Pages);
        OfficeDocumentModelBlock block = Assert.Single(model.Blocks);
        string[] adapterReferences = typeof(VisioPdfConverterExtensions).Assembly
            .GetReferencedAssemblies()
            .Select(static assembly => assembly.Name ?? string.Empty)
            .ToArray();

        Assert.Equal(OfficeDocumentFormat.Visio, model.Format);
        Assert.Equal("topology.vsdx", model.Source.Path);
        Assert.Equal("Topology", page.Name);
        Assert.Contains("Gateway service", block.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(adapterReferences, static name =>
            name.StartsWith("OfficeIMO.Reader", StringComparison.Ordinal));
    }

    [Fact]
    public void LoadedDocumentsUseTheirAssociatedPathsAsDistinctReaderIdentities() {
        string firstPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".vsdx");
        string secondPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".vsdx");
        try {
            CreateDiagram().Save(firstPath);
            CreateDiagram().Save(secondPath);
            VisioDocument first = VisioDocument.Load(firstPath);
            VisioDocument second = VisioDocument.Load(secondPath);

            OfficeDocumentReadResult firstResult = first.ToOfficeDocumentReadResult();
            OfficeDocumentReadResult secondResult = second.ToOfficeDocumentReadResult();

            Assert.Equal(Path.GetFullPath(firstPath), firstResult.Source.Path);
            Assert.Equal(Path.GetFullPath(secondPath), secondResult.Source.Path);
            Assert.NotEqual(firstResult.Source.SourceId, secondResult.Source.SourceId);
        } finally {
            if (File.Exists(firstPath)) File.Delete(firstPath);
            if (File.Exists(secondPath)) File.Delete(secondPath);
        }
    }

    [Fact]
    public async Task PdfLifecycleSupportsBytesPathStreamAndCancellation() {
        VisioDocument diagram = CreateDiagram();
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            byte[] bytes = diagram.ToPdf();
            PdfSaveResult saved = diagram.SaveAsPdf(outputPath);
            using var stream = new MemoryStream();
            PdfSaveResult streamed = await diagram.SaveAsPdfAsync(stream);

            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
            Assert.True(saved.Succeeded);
            Assert.True(File.Exists(outputPath));
            Assert.True(streamed.Succeeded);
            Assert.True(stream.CanWrite);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(stream.ToArray(), 0, 4));

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                diagram.SaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
        } finally {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void NeutralProjectionPropagatesOperationCancellationIntoPngPreviewRendering() {
        VisioDocument diagram = CreateDiagram();
        using var cancellation = new CancellationTokenSource();
        var provider = new CancelingTextShapingProvider(cancellation);
        var pngOptions = new VisioPngSaveOptions { TextShapingProvider = provider };

        Assert.Throws<OperationCanceledException>(() => diagram.ToOfficeDocumentModel(
            options: new VisioDocumentProjectionOptions {
                IncludePngPreviewAssets = true,
                PngOptions = pngOptions
            },
            cancellationToken: cancellation.Token));
        Assert.True(provider.WasCalled);
        Assert.False(pngOptions.CancellationToken.IsCancellationRequested);
    }

    [Fact]
    public void PdfAdapterMatchesTheSharedDocumentLifecycle() {
        MethodInfo[] methods = typeof(VisioPdfConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static);

        Assert.Single(methods, method => method.Name == "ToPdf");
        Assert.Single(methods, method => method.Name == "ToPdfDocument");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentResult");
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfAsync"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdfAsync"));
    }

    private static VisioDocument CreateDiagram() {
        VisioDocument diagram = VisioDocument.Create();
        VisioPage page = diagram.AddPage("Topology", 8, 5);
        page.Shapes.Add(new VisioShape("gateway") { Text = "Gateway service" });
        return diagram;
    }

    private sealed class CancelingTextShapingProvider : IOfficeTextShapingProvider {
        private readonly CancellationTokenSource _cancellation;

        internal CancelingTextShapingProvider(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        internal bool WasCalled { get; private set; }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            WasCalled = true;
            _cancellation.Cancel();
            return null;
        }
    }
}
