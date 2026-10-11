using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public class WordImageWorkflowTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task MetadataDiagnosticsDistinguishProposalsFromAppliedRemoval(bool analyze, bool allowLoss) {
        await InDirectory(async root => {
            string input = Path.Combine(root, "source.docx"), output = Path.Combine(root, "output.docx");
            byte[] source = CreateWord(withMetadata: true);
            File.WriteAllBytes(input, source);
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = analyze ? OfficeWorkflowOperation.AnalyzeWordImages : OfficeWorkflowOperation.OptimizeWordImages,
                InputPath = input, OutputPath = analyze ? null : output,
                WordImageOptimization = new() { MetadataPolicy = OfficeImageMetadataPolicy.Strip, AllowMetadataLoss = allowLoss }
            });
            Assert.True(result.Succeeded, result.Summary);
            var diagnostic = Assert.Single(result.Diagnostics, item => item.Code == (allowLoss ? "WordImageOptimized" : "WordImageMetadataLoss"));
            Assert.Equal(OfficeWorkflowDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Contains("Xmp", diagnostic.Details["candidateMetadataStripped"]);
            Assert.Equal((allowLoss && !analyze).ToString(), diagnostic.Details["applied"]);
            Assert.Equal("400", diagnostic.Details["originalWidth"]);
            var inventory = Assert.Single(result.Diagnostics, item => item.Code == "WordImageInventory");
            Assert.Equal(allowLoss, long.Parse(inventory.Details["requiredStagedBytes"]) > 0);
            Assert.Equal(source, File.ReadAllBytes(input));
            if (!analyze) {
                using WordDocument saved = WordDocument.Load(output);
                string image = System.Text.Encoding.UTF8.GetString(saved.Images[0].ToBytes());
                Assert.Equal(!allowLoss, image.Contains("private image metadata"));
            }
        });
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("doc")]
    [InlineData("pdf")]
    public async Task PublishesReopenedCopyAndProtectsInput(string format) {
        await InDirectory(async root => {
            string input = Path.Combine(root, "source.docx"), output = Path.Combine(root, "optimized." + format);
            byte[] original = CreateWord();
            File.WriteAllBytes(input, original);
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.OptimizeWordImages, InputPath = input, OutputPath = output,
                WordImageOptimization = new() { Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 25 }
            });
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(original, File.ReadAllBytes(input));
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "WordImageOptimized");
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
            Assert.Equal(2, Directory.GetFiles(root).Length);
            if (format != "pdf") {
                using WordDocument reopened = WordDocument.Load(output);
                Assert.Equal(400, OfficeImageReader.Identify(reopened.Images[0].ToBytes()).Width);
            } else {
                Assert.Single(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(File.ReadAllBytes(output)));
                var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
                foreach (var finding in evidence.FidelityDiagnostics)
                    Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == finding.Code && diagnostic.Message == finding.Message
                        && diagnostic.Details["source"] == finding.Source
                        && diagnostic.Details["lossKind"] == finding.LossKind.ToString());
            }
        });
    }

    [Fact]
    public async Task AnalyzeAndFailedPublicationRetainSourceAndExistingDestination() {
        await InDirectory(async root => {
            string input = Path.Combine(root, "source.docx"), output = Path.Combine(root, "optimized.docx");
            byte[] original = CreateWord();
            File.WriteAllBytes(input, original);
            File.WriteAllText(output, "existing destination");
            var runner = new OfficeWorkflowRunner();
            var analysis = await runner.RunAsync(new() { Operation = OfficeWorkflowOperation.AnalyzeWordImages, InputPath = input });
            Assert.True(analysis.Succeeded, analysis.Summary);
            Assert.Null(analysis.OutputPath);
            var refused = await runner.RunAsync(new() { Operation = OfficeWorkflowOperation.OptimizeWordImages, InputPath = input, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Fail });
            Assert.False(refused.Succeeded);
            Assert.Equal("existing destination", File.ReadAllText(output));
            var sourceOverwrite = await runner.RunAsync(new() { Operation = OfficeWorkflowOperation.OptimizeWordImages, InputPath = input, OutputPath = input, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace });
            Assert.False(sourceOverwrite.Succeeded);
            Assert.Equal(original, File.ReadAllBytes(input));
            Assert.Equal(2, Directory.GetFiles(root).Length);
        });
    }

    [Fact]
    public async Task BatchSnapshotsOptionsAndContinuesAfterInvalidInput() {
        await InDirectory(async root => {
            string input = Path.Combine(root, "source.docx"), output = Path.Combine(root, "output.docx");
            File.WriteAllBytes(input, CreateWord());
            var options = new WordImageOptimizationOptions { Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 25 };
            var progress = new InlineProgress(_ => options.Mode = (OfficeImageOptimizationMode)99);
            var results = await new OfficeWorkflowRunner().RunBatchAsync([
                new() { Operation = OfficeWorkflowOperation.OptimizeWordImages, InputPath = Path.Combine(root, "missing.docx"), OutputPath = output },
                new() { Operation = OfficeWorkflowOperation.OptimizeWordImages, InputPath = input, OutputPath = output, WordImageOptimization = options }
            ], progress);
            Assert.False(results[0].Succeeded);
            Assert.True(results[1].Succeeded, results[1].Summary);
            Assert.Contains(results[1].Diagnostics, item => item.Code == "WordImageOptimized");
        });
    }

    private static byte[] CreateWord(bool withMetadata = false) {
        using WordDocument word = WordDocument.Create();
        var raster = new OfficeRasterImage(400, 200);
        for (int y = 0; y < 200; y++) for (int x = 0; x < 400; x++) raster.SetPixel(x, y, OfficeColor.FromRgb((byte)(x * 17), (byte)(y * 29), (byte)(x * y)));
        using var image = new MemoryStream(OfficeJpegCodec.Encode(raster, new() {
            Quality = 98, Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = new OfficeJpegMetadata(xmp: withMetadata ? System.Text.Encoding.UTF8.GetBytes("private image metadata") : null)
        }));
        word.AddParagraph("Editable text").InsertImage(image, "image.jpg", 96, 48);
        return word.ToBytes();
    }
    private static async Task InDirectory(Func<string, Task> action) {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-WordImageTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try { await action(root); } finally { Directory.Delete(root, recursive: true); }
    }
    private sealed class InlineProgress(Action<OfficeWorkflowProgress> action) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => action(value);
    }
}
