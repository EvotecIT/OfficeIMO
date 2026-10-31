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

    private sealed class RejectingStream : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("Destination rejected the bytes.");
        public override void Write(ReadOnlySpan<byte> buffer) => throw new IOException("Destination rejected the bytes.");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => Task.FromException(new IOException("Destination rejected the bytes."));
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) => new(Task.FromException(new IOException("Destination rejected the bytes.")));
    }
}
