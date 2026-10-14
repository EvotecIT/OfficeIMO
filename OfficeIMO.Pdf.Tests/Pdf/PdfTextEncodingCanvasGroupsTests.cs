using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextEncodingCanvasGroupsTests {
    [Theory]
    [InlineData("structure")]
    [InlineData("actual-text")]
    [InlineData("artifact")]
    [InlineData("figure")]
    public void AnalyzeTextEncoding_TraversesNestedCanvasPaintWithoutTreatingLogicalUnicodeAsGlyphs(string group) {
        PdfDocument document = PdfDocument.Create().Canvas(canvas => {
            void Paint(PdfPageCanvas child) => child.Text("A−B", 0D, 0D, 100D, 20D);
            switch (group) {
                case "structure": canvas.Structure(PdfCanvasStructureRole.Paragraph, Paint); break;
                case "actual-text": canvas.ActualText("Logical Ω", Paint); break;
                case "artifact": canvas.Artifact(Paint); break;
                case "figure": canvas.Figure("Alternative Ω", Paint); break;
            }
        });

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());

        Assert.Equal("U+2212", diagnostic.CodePoint);
        Assert.Equal("PdfCanvasText", diagnostic.Source);
        Assert.Contains("PdfCanvas", diagnostic.Location, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
    }

    [Fact]
    public void AnalyzeTextEncoding_CanceledTokenStopsBeforeNestedCanvasTraversal() {
        PdfDocument document = PdfDocument.Create().Canvas(canvas =>
            canvas.Structure(PdfCanvasStructureRole.Paragraph, child =>
                child.ActualText("Logical Ω", paint => paint.Text("A−B", 0D, 0D, 100D, 20D))));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OperationCanceledException exception = Assert.Throws<OperationCanceledException>(() =>
            document.AnalyzeTextEncoding(cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        PdfTextEncodingDiagnostic original = Assert.Single(document.AnalyzeTextEncoding());
        PdfTextEncodingDiagnostic uncanceled = Assert.Single(document.AnalyzeTextEncoding(CancellationToken.None));
        Assert.Equal(original.CodePoint, uncanceled.CodePoint);
        Assert.Equal(original.Source, uncanceled.Source);
        Assert.Equal(original.Location, uncanceled.Location);
    }

    [Fact]
    public void AnalyzeTextEncoding_CancellationDuringDeferredRowsStopsAndDisposesEnumeration() {
        using var cancellation = new CancellationTokenSource();
        int yieldedRows = 0;
        bool disposed = false;
        IEnumerable<string[]> CreateRows() {
            try {
                yieldedRows++;
                yield return new[] { "First row" };
                cancellation.Cancel();
                yieldedRows++;
                yield return new[] { "A−B" };
                yieldedRows++;
                yield return new[] { "Unvisited row" };
            } finally {
                disposed = true;
            }
        }
        PdfDocument document = PdfDocument.Create();
        document.Compose(builder => builder.Content(content => content.TableDeferred(CreateRows, batchSize: 1)));

        OperationCanceledException exception = Assert.Throws<OperationCanceledException>(() =>
            document.AnalyzeTextEncoding(cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(2, yieldedRows);
        Assert.True(disposed);
    }
}
