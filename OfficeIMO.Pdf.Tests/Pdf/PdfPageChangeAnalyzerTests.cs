using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageChangeAnalyzerTests {
    [Fact]
    public void AlignsInsertedModifiedAndMovedPagesWithoutTreatingInsertionAsMovement() {
        PdfDocument expected = BuildPages("Alpha", "Bravo", "Charlie");
        PdfDocument inserted = BuildPages("Intro", "Alpha", "Bravo", "Charlie");
        PdfPageChangeReport insertion = expected.Proof.AnalyzePageChanges(inserted);

        Assert.Equal(new[] { PdfPageChangeKind.Unchanged, PdfPageChangeKind.Unchanged, PdfPageChangeKind.Unchanged, PdfPageChangeKind.Inserted },
            insertion.Changes.Select(static change => change.Kind));
        Assert.Equal(4, insertion.Changes[2].ActualPageNumber);
        Assert.Equal(1, insertion.Changes[3].ActualPageNumber);

        PdfPageChangeReport modified = expected.Proof.AnalyzePageChanges(BuildPages("Alpha", "Changed", "Charlie"));
        Assert.Equal(PdfPageChangeKind.ModifiedCandidate, modified.Changes[1].Kind);
        Assert.Equal(2, modified.Changes[1].ActualPageNumber);

        PdfPageChangeReport moved = expected.Proof.AnalyzePageChanges(BuildPages("Bravo", "Charlie", "Alpha"));
        Assert.Single(moved.Changes, static change => change.Kind == PdfPageChangeKind.Moved);
        Assert.All(moved.Changes, static change => Assert.True(change.IsExactRenderedMatch));

        PdfVisualPageComparison aligned = expected.Proof.CompareVisualPages(1, inserted, 2);
        Assert.True(aligned.IsMatch);
        Assert.Equal(1, aligned.PageNumber);
        Assert.Equal(2, aligned.ActualPageNumber);
        PdfVisualPageComparison changed = expected.Proof.CompareVisualPages(2, BuildPages("Alpha", "Changed", "Charlie"), 2);
        Assert.False(changed.IsMatch);
        Assert.NotNull(changed.ChangedBounds);
    }

    [Fact]
    public void PageAlignmentHonorsPixelBudgetAndCancellation() {
        PdfDocument source = BuildPages("One");
        Assert.Throws<PdfReadLimitException>(() => source.Proof.AnalyzePageChanges(source,
            new PdfPageChangeOptions { MaxPixelsPerPage = 1 }));
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => source.Proof.AnalyzePageChanges(source, cancellationToken: canceled.Token));
    }

    private static PdfDocument BuildPages(params string[] texts) {
        PdfDocument document = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) });
        for (int index = 0; index < texts.Length; index++) {
            if (index != 0) document.PageBreak();
            document.Paragraph(paragraph => paragraph.Text(texts[index]));
        }
        return PdfDocument.Load(document.ToBytes());
    }
}
