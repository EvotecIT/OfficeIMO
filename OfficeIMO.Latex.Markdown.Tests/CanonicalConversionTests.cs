using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class CanonicalConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SavesRetainDestinationAndNativeEvidence(bool asynchronous) {
        MarkdownDocument source = MarkdownDocument.Parse("# Title\n\n**Bold** and *italic*.\n\n---\n");
        var options = new MarkdownLatexExportOptions { LineEnding = "\r\n" };
        OfficeConversionResult<LatexDocument> converted = await source.ToLatexAsync(options);
        MarkdownToLatexReport report = Assert.IsType<MarkdownToLatexReport>(converted.Report);
        Assert.Null(report.OutputPath);
        Assert.Contains("Title", report.GeneratedSource, StringComparison.Ordinal);
        Assert.NotEmpty(report.NativeDiagnostics);
        using var stream = new MemoryStream();
        OfficeConversionReport saved = asynchronous
            ? await source.SaveAsLatexAsync(stream, options)
            : source.SaveAsLatex(stream, options);
        Assert.Null(saved.OutputPath);
        Assert.Equal(report.GeneratedSource, Encoding.UTF8.GetString(stream.ToArray()));
        Assert.Equal(report.GeneratedSource, Assert.IsType<MarkdownToLatexReport>(saved).GeneratedSource);
        Assert.True(stream.CanWrite);
        OfficeConversionResult<MarkdownDocument> back = await converted.Value.ToMarkdownAsync();
        Assert.Contains("**Bold**", back.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.IsType<LatexToMarkdownReport>(back.Report);
        Assert.Null(back.Report.OutputPath);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-latex-save-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "source.tex");
            string relativePath = Path.Combine(directory, ".", "source.tex");
            OfficeConversionReport file = asynchronous
                ? await source.SaveAsLatexAsync(relativePath, options)
                : source.SaveAsLatex(relativePath, options);
            Assert.Equal(Path.GetFullPath(path), file.OutputPath);
            MarkdownToLatexReport native = Assert.IsType<MarkdownToLatexReport>(file);
            Assert.Equal(report.GeneratedSource, native.GeneratedSource);
            Assert.Equal(native.GeneratedSource, File.ReadAllText(path));
            Assert.Equal(report.NativeDiagnostics.Select(d => (d.Code, d.Outcome)),
                native.NativeDiagnostics.Select(d => (d.Code, d.Outcome)));
            Assert.Equal(report.Diagnostics.Select(d => (d.Code, d.LossKind)),
                file.Diagnostics.Select(d => (d.Code, d.LossKind)));
            Assert.Throws<OfficeConversionException>(() => file.ThrowIfLoss());

            LatexDocument reopened = LatexDocument.Parse(File.ReadAllText(path));
            string markdownPath = Path.Combine(directory, "back.md");
            OfficeConversionReport markdown = asynchronous
                ? await reopened.SaveAsMarkdownAsync(markdownPath)
                : reopened.SaveAsMarkdown(markdownPath);
            Assert.Equal(Path.GetFullPath(markdownPath), markdown.OutputPath);
            Assert.IsType<LatexToMarkdownReport>(markdown);
            Assert.Contains("**Bold**", MarkdownDocument.Parse(File.ReadAllText(markdownPath)).ToMarkdown(), StringComparison.Ordinal);
            using var markdownStream = new MemoryStream();
            OfficeConversionReport streamReport = asynchronous
                ? await reopened.SaveAsMarkdownAsync(markdownStream)
                : reopened.SaveAsMarkdown(markdownStream);
            Assert.Null(streamReport.OutputPath);
            Assert.IsType<LatexToMarkdownReport>(streamReport);
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
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.SaveAsLatexAsync(stream, cancellationToken: cancelled.Token));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
}
