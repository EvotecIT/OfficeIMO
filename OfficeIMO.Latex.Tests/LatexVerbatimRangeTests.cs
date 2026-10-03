using System.Threading;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexVerbatimRangeTests {
    [Theory]
    [InlineData(4093)]
    [InlineData(4094)]
    [InlineData(4095)]
    [InlineData(4096)]
    public void CancellableOpaqueSearchPreservesClosingDelimiterAcrossSearchChunks(int padding) {
        string source = "\\begin{document}\\begin{verbatim}" + new string('x', padding) + "\\end{verbatim}After\\end{document}";
        using var cancellation = new CancellationTokenSource();
        LatexDocument document = LatexDocument.Parse(source, null, cancellation.Token);
        Assert.Equal(source, document.ToLatex());
        Assert.Contains(document.Paragraphs, static paragraph => paragraph.Content == "After");
        Assert.DoesNotContain(document.Diagnostics, static diagnostic => diagnostic.Severity == LatexDiagnosticSeverity.Error);
    }
}
