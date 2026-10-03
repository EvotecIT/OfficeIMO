using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexPdfAdapterContractTests {
    private const string MinimalDocument = "\\documentclass{article}\n\\begin{document}\nLifecycle marker.\n\\end{document}\n";

    [Theory]
    [InlineData("100\\% useful", "A\\\\", "100% useful")]
    [InlineData("\\textbf{R\\&D} 50\\%", "A\\\\", "R&D 50%")]
    [InlineData("Only 100\\%", "", "Only 100%")]
    public void TableCaptionsRemainVisibleInRenderedPdf(string caption, string rows, string expected) {
        LatexDocument document = LatexDocument.Parse("\\documentclass{article}\\begin{document}\\begin{table}\\caption{" + caption +
            "}\\begin{tabular}{l}" + rows + "\\end{tabular}\\end{table}\\end{document}");
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfDocumentResult().Value.ToBytes()).ExtractText();
        Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.DoesNotContain("\\%", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\\textbf", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ParserDiagnostics_FlowIntoFinalPdfResult() {
        LatexDocument document = LatexDocument.ParseResult("\\documentclass{article}\n\\begin{document}\nBroken $x^2\n").Document;

        var result = document.ToPdfDocumentResult();
        var warning = Assert.Single(result.Warnings, item => item.Code == "LATEX003");

        Assert.Equal("OfficeIMO.Latex.Pdf", warning.Converter);
        Assert.Equal(OfficeIMO.Pdf.PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("parse", warning.Details["stage"]);
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void SaveAsPdf_LeavesCallerOwnedStreamOpen() {
        LatexDocument document = LatexDocument.ParseResult(MinimalDocument).Document;
        using var stream = new MemoryStream();

        document.SaveAsPdf(stream);
        long written = stream.Length;
        stream.Position = stream.Length;
        stream.WriteByte(0);

        Assert.True(written > 100);
        Assert.Equal(written + 1, stream.Length);
    }

    [Fact]
    public async Task SaveAsPdfAsync_PreCanceled_DoesNotConvertOrWrite() {
        LatexDocument document = LatexDocument.ParseResult(MinimalDocument).Document;
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsPdfAsync(stream, cancellationToken: cancellation.Token));

        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
}
