using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConversionPublicationFailureTests {
    [Theory]
    [InlineData("Reader", false)]
    [InlineData("Reader", true)]
    [InlineData("Rtf", false)]
    [InlineData("Rtf", true)]
    [InlineData("ImagePaths", false)]
    [InlineData("ImagePaths", true)]
    [InlineData("ImageBytes", false)]
    [InlineData("ImageBytes", true)]
    public async Task NormalPdfSavesRetainReportsOnDestinationFailure(string route, bool asynchronous) {
        string directory = Directory.CreateTempSubdirectory("officeimo-save-failure-").FullName;
        try {
            using var output = new RejectingStream();
            OfficeConversionException error = asynchronous
                ? await Assert.ThrowsAsync<OfficeConversionException>(Save)
                : Assert.Throws<OfficeConversionException>(() => Save().GetAwaiter().GetResult());
            var report = Assert.IsType<OfficeConversionReport>(error.Report);
            Assert.Equal(route == "Reader" ? "Word" : route.StartsWith("Image", StringComparison.Ordinal) ? "Image" : "Rtf", report.Source);
            Assert.Equal("Pdf", report.Target);
            Assert.Contains(report.Diagnostics, d => d.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
            Assert.IsType<IOException>(error.InnerException);
            Assert.True(output.CanWrite);

            async Task Save() {
                if (route == "Reader") {
                    using var word = WordDocument.Create(); word.AddParagraph("Reader publication evidence");
                    using var input = new MemoryStream(word.ToBytes());
                    var reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
                    var source = reader.ReadDocument(input, "source.docx");
                    if (asynchronous) await source.SaveAsPdfAsync(output);
                    else source.SaveAsPdf(output);
                } else if (route == "Rtf") {
                    var source = RtfDocument.Create(); source.AddSection().AddParagraph("RTF publication evidence");
                    if (asynchronous) await source.SaveAsPdfAsync(output);
                    else source.SaveAsPdf(output);
                } else {
                    byte[] bytes = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
                    if (route == "ImagePaths") {
                        string input = Path.Combine(directory, "image.png"); File.WriteAllBytes(input, bytes);
                        string[] source = [input];
                        if (asynchronous) await source.SaveAsPdfAsync(output);
                        else source.SaveAsPdf(output);
                    } else {
                        PdfImageDocumentSource[] source = [new(bytes, "image.png")];
                        if (asynchronous) await source.SaveAsPdfAsync(output);
                        else source.SaveAsPdf(output);
                    }
                }
            }
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(OfficePublicationPolicy.Fail)]
    [InlineData(OfficePublicationPolicy.Rename)]
    public async Task ConcurrentPublicationsKeepEachSuccessfulArtifactsBytes(OfficePublicationPolicy policy) {
        string directory = Directory.CreateTempSubdirectory("officeimo-concurrent-publication-").FullName;
        try {
            const int writers = 16;
            string destination = Path.Combine(directory, "output.txt");
            var guard = new PublicationRendezvous(destination, writers);
            var operations = Enumerable.Range(0, writers).Select(async index => {
                string payload = "Complete writer " + index;
                string stage = Path.Combine(directory, index + ".tmp");
                File.WriteAllText(stage, payload);
                try {
                    string published = await OfficeFilePublisher.PublishAsync(stage, destination, policy, guard, CancellationToken.None);
                    return (Path: published, Payload: payload);
                } catch (IOException) when (policy == OfficePublicationPolicy.Fail) {
                    return (Path: (string?)null, Payload: payload);
                }
            }).ToArray();
            var results = (await Task.WhenAll(operations)).Where(item => item.Path != null).ToArray();
            Assert.Equal(policy == OfficePublicationPolicy.Fail ? 1 : writers, results.Length);
            Assert.Equal(results.Length, results.Select(item => item.Path).Distinct().Count());
            foreach (var item in results) Assert.Equal(item.Payload, File.ReadAllText(item.Path!));
        } finally { Directory.Delete(directory, true); }
    }

    private sealed class PublicationRendezvous(string destination, int writers) : IOfficePublicationGuard {
        private readonly TaskCompletionSource<bool> _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _arrived;

        public async Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) {
            if (absoluteDestination == destination) {
                if (Interlocked.Increment(ref _arrived) == writers) _ready.TrySetResult(true);
                if (await Task.WhenAny(_ready.Task, Task.Delay(TimeSpan.FromSeconds(30), cancellationToken)) != _ready.Task)
                    throw new TimeoutException("Concurrent publication writers did not reach their completed staging files.");
                await _ready.Task;
            }
            return true;
        }
    }

    private sealed class RejectingStream : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("Destination rejected the bytes.");
        public override void Write(ReadOnlySpan<byte> buffer) => throw new IOException("Destination rejected the bytes.");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => Task.FromException(new IOException("Destination rejected the bytes."));
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) => new(Task.FromException(new IOException("Destination rejected the bytes.")));
    }
}
