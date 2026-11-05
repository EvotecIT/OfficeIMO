using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class CanonicalConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SavesRetainDestinationAndNativeEvidence(bool asynchronous) {
        MarkdownDocument source = MarkdownDocument.Parse("# Title\n\n**Bold** and *italic*.\n\n---\n");
        var options = new MarkdownAsciiDocExportOptions { LineEnding = "\r\n" };
        OfficeConversionResult<AsciiDocDocument> converted = await source.ToAsciiDocAsync(options);
        MarkdownToAsciiDocReport report = Assert.IsType<MarkdownToAsciiDocReport>(converted.Report);
        Assert.Null(report.OutputPath);
        Assert.Contains("Title", report.GeneratedSource, StringComparison.Ordinal);
        Assert.NotEmpty(report.NativeDiagnostics);
        using var stream = new MemoryStream();
        OfficeConversionReport saved = asynchronous
            ? await source.SaveAsAsciiDocAsync(stream, options)
            : source.SaveAsAsciiDoc(stream, options);
        Assert.Null(saved.OutputPath);
        Assert.Equal(report.GeneratedSource, Encoding.UTF8.GetString(stream.ToArray()));
        Assert.Equal(report.GeneratedSource, Assert.IsType<MarkdownToAsciiDocReport>(saved).GeneratedSource);
        Assert.True(stream.CanWrite);
        OfficeConversionResult<MarkdownDocument> back = await converted.Value.ToMarkdownAsync();
        Assert.Contains("**Bold**", back.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.IsType<AsciiDocToMarkdownReport>(back.Report);
        Assert.Null(back.Report.OutputPath);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-asciidoc-save-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "source.adoc");
            string relativePath = Path.Combine(directory, ".", "source.adoc");
            OfficeConversionReport file = asynchronous
                ? await source.SaveAsAsciiDocAsync(relativePath, options)
                : source.SaveAsAsciiDoc(relativePath, options);
            Assert.Equal(Path.GetFullPath(path), file.OutputPath);
            MarkdownToAsciiDocReport native = Assert.IsType<MarkdownToAsciiDocReport>(file);
            Assert.Equal(report.GeneratedSource, native.GeneratedSource);
            Assert.Equal(native.GeneratedSource, File.ReadAllText(path));
            Assert.Equal(report.NativeDiagnostics.Select(d => (d.Code, d.Outcome)),
                native.NativeDiagnostics.Select(d => (d.Code, d.Outcome)));
            Assert.Equal(report.Diagnostics.Select(d => (d.Code, d.LossKind)),
                file.Diagnostics.Select(d => (d.Code, d.LossKind)));
            Assert.Throws<OfficeConversionException>(() => file.ThrowIfLoss());

            AsciiDocDocument reopened = AsciiDocDocument.Parse(File.ReadAllText(path));
            string markdownPath = Path.Combine(directory, "back.md");
            OfficeConversionReport markdown = asynchronous
                ? await reopened.SaveAsMarkdownAsync(markdownPath)
                : reopened.SaveAsMarkdown(markdownPath);
            Assert.Equal(Path.GetFullPath(markdownPath), markdown.OutputPath);
            Assert.IsType<AsciiDocToMarkdownReport>(markdown);
            Assert.Contains("**Bold**", MarkdownDocument.Parse(File.ReadAllText(markdownPath)).ToMarkdown(), StringComparison.Ordinal);
            using var markdownStream = new MemoryStream();
            OfficeConversionReport streamReport = asynchronous
                ? await reopened.SaveAsMarkdownAsync(markdownStream)
                : reopened.SaveAsMarkdown(markdownStream);
            Assert.Null(streamReport.OutputPath);
            Assert.IsType<AsciiDocToMarkdownReport>(streamReport);
            Assert.Equal(File.ReadAllText(markdownPath), Encoding.UTF8.GetString(markdownStream.ToArray()));
            Assert.True(markdownStream.CanWrite);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task CancelledConversionDoesNotWriteTheCallerStream() {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        using var stream = new MemoryStream();
        MarkdownDocument source = MarkdownDocument.Parse("Visible");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.SaveAsAsciiDocAsync(stream, cancellationToken: cancelled.Token));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
    [Fact]
    public void StrictSourceFallbackPreservesEvidenceBeforeWriting() {
        MarkdownDocument source = MarkdownDocument.Create().Hr();
        using var stream = new MemoryStream();
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() =>
            source.SaveAsAsciiDoc(stream, new MarkdownAsciiDocExportOptions { FailOnLoss = true }));
        MarkdownToAsciiDocReport report = Assert.IsType<MarkdownToAsciiDocReport>(error.Report);
        Assert.Contains("[source,markdown]", report.GeneratedSource, StringComparison.Ordinal);
        Assert.True(report.HasLoss);
        Assert.Equal(0, stream.Length);
    }
}
