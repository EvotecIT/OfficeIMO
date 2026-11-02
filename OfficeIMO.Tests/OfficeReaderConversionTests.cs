using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeReaderConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NormalizedContentUsesTheSameProjectionAndKeepsCallerStreamsOpen(bool asynchronous) {
        using var word = WordDocument.Create();
        word.AddParagraph("Normalized source content");
        using var input = new MemoryStream(word.ToBytes());
        var reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        OfficeDocumentReadResult source = reader.ReadDocument(input, "source.docx");
        var options = new PdfProjectionOptions { IncludeMetadata = false };
        using PdfDocument model = asynchronous
            ? (await source.ToPdfAsync(options)).Value
            : source.ToPdf(options, out _);
        Assert.Contains("Normalized source content", model.ExtractText());
        using var destination = new MemoryStream();
        OfficeConversionReport report = asynchronous
            ? await source.SaveAsPdfAsync(destination, options)
            : source.SaveAsPdf(destination, options);
        Assert.Equal("Pdf", report.Target);
        Assert.True(report.ItemsWritten > 0);
        Assert.True(destination.CanWrite);
        destination.Position = 0;
        using var loaded = PdfDocument.Load(destination);
        Assert.Contains("Normalized source content", loaded.ExtractText());
        Assert.Contains("Normalized source content", loaded.ExtractMarkdown());
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        byte[] retained = destination.ToArray();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.SaveAsPdfAsync(destination, options, cancelled.Token));
        Assert.Equal(retained, destination.ToArray());
    }
}
