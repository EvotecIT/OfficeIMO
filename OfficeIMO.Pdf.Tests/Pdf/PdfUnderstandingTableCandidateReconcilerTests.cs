using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfUnderstandingTableCandidateReconcilerTests {
    [Fact]
    public void Reconcile_ChargesCandidatePreparationToTheUnderstandingBudget() {
        PdfLogicalPage page = CreatePage();
        var budget = new PdfUnderstandingWorkBudget(12, CancellationToken.None);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfUnderstandingTableCandidateReconciler.Reconcile(
                page,
                new[] { CreateCandidate(0) },
                new[] { CreateCandidate(1) },
                budget.Consume,
                budget.ThrowIfCancellationRequested));

        Assert.Equal(PdfReadLimitKind.UnderstandingWork, exception.Kind);
        Assert.Equal(12, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void Reconcile_ObservesCancellationDuringPairwiseComparison() {
        PdfLogicalPage page = CreatePage();
        PdfUnderstandingTableCandidate[] candidates = Enumerable.Range(1, 80)
            .Select(CreateCandidate)
            .ToArray();
        using var cancellation = new CancellationTokenSource();
        long consumed = 0;

        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfUnderstandingTableCandidateReconciler.Reconcile(
                page,
                new[] { CreateCandidate(0) },
                candidates,
                units => {
                    consumed += units;
                    if (consumed >= 250) cancellation.Cancel();
                },
                cancellation.Token.ThrowIfCancellationRequested);
        });

        Assert.True(consumed >= 250);
    }

    [Fact]
    public void Reconcile_KeepsPartialTaggedOwnershipSeparateFromMixedSourceWords() {
        PdfLogicalPage page = CreatePage();
        var taggedRun = new PdfTextSpan("Tagged", "F1", 10D, 20D, 700D, 40D);
        var untaggedRun = new PdfTextSpan("Untagged", "F1", 10D, 80D, 700D, 50D);
        PdfUnderstandingLine mixedLine = CreateLine("Tagged Untagged", taggedRun, untaggedRun);
        PdfUnderstandingTableCandidate geometry = CreateCandidate(
            "geometry",
            new[] { mixedLine },
            new[] { taggedRun, untaggedRun });
        PdfUnderstandingTableCandidate tagged = CreateTaggedCandidate(
            "Tagged",
            new[] { mixedLine },
            new[] { taggedRun });

        IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
            PdfUnderstandingTableCandidateReconciler.Reconcile(page, new[] { geometry }, new[] { tagged });

        Assert.Equal(new[] { taggedRun }, tagged.NativeSourceRuns);
        Assert.Equal(2, reconciled.Count);
        Assert.Contains(geometry, reconciled);
        Assert.Contains(tagged, reconciled);
    }

    [Fact]
    public void TaggedTableProjectionDoesNotInheritUntaggedRunsFromCustomMergedWords() {
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Table, table => table
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Code", 40D, 100D, 60D, 16D))
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Value", 140D, 100D, 60D, 16D)))
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("A-01", 40D, 120D, 60D, 16D))
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("10", 140D, 120D, 60D, 16D))))
                .Text("Outside header", 260D, 100D, 100D, 16D)
                .Text("Outside value", 260D, 120D, 100D, 16D))
            .ToBytes();
        PdfUnderstandingPipelineOptions pipeline = PdfUnderstandingPipelineOptions.Structured();
        pipeline.WordGrouping = new BaselineMergedWordGroupingStage();

        PdfLogicalPage page = Assert.Single(PdfDocument.Load(pdf).Read(new PdfReadOptions {
            Profile = PdfReadProfile.Structured,
            Pipeline = pipeline
        }).Pages);
        PdfUnderstandingTableCandidate tagged = Assert.Single(
            page.Analysis.TableCandidates,
            static candidate => candidate.DetectionKind == "tagged-structure");

        Assert.Equal(4, tagged.NativeSourceRuns.Count);
        Assert.All(tagged.NativeSourceRuns, static run => Assert.True(run.MarkedContentId.HasValue));
        Assert.DoesNotContain(tagged.NativeSourceRuns, static run => run.Text.Contains("Outside", StringComparison.Ordinal));
        Assert.All(
            tagged.SourceLines.SelectMany(static line => line.Words).SelectMany(static word => word.SourceRuns),
            static run => Assert.True(run.MarkedContentId.HasValue));
    }

    [Fact]
    public void TaggedTableUnionFullyAccountsForOneSpanningGeometricCandidate() {
        var firstRun = new PdfTextSpan("A", "F1", 10D, 20D, 700D, 10D);
        var secondRun = new PdfTextSpan("1", "F1", 10D, 80D, 700D, 10D);
        var thirdRun = new PdfTextSpan("B", "F1", 10D, 20D, 660D, 10D);
        var fourthRun = new PdfTextSpan("2", "F1", 10D, 80D, 660D, 10D);
        PdfUnderstandingLine firstLine = CreateLine("A 1", firstRun, secondRun);
        PdfUnderstandingLine secondLine = CreateLine("B 2", thirdRun, fourthRun);
        PdfUnderstandingTableCandidate geometry = CreateCandidate(
            "geometry",
            new[] { firstLine, secondLine },
            new[] { firstRun, secondRun, thirdRun, fourthRun });
        PdfUnderstandingTableCandidate firstTagged = CreateTaggedCandidate(
            "A",
            new[] { firstLine },
            new[] { firstRun, secondRun });
        PdfUnderstandingTableCandidate secondTagged = CreateTaggedCandidate(
            "B",
            new[] { secondLine },
            new[] { thirdRun, fourthRun });
        var budget = new PdfUnderstandingWorkBudget(10_000, CancellationToken.None);

        IReadOnlyList<PdfUnderstandingTableCandidate> retained =
            PdfDocumentSemanticEnricher.RemoveFullyAccountedGeometricTables(
                new[] { geometry },
                new List<PdfUnderstandingTableCandidate> { firstTagged, secondTagged },
                new HashSet<PdfTextSpan>(),
                budget);

        Assert.Empty(retained);
    }

    private static PdfLogicalPage CreatePage() => Assert.Single(
        PdfDocumentReadResult.Load(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("anchor")).ToBytes()).Pages);

    private static PdfUnderstandingTableCandidate CreateCandidate(int index) {
        double left = index * 220D;
        return new PdfUnderstandingTableCandidate(
            "test-geometry",
            700D,
            680D,
            new[] {
                new PdfUnderstandingTableColumn(left, left + 80D),
                new PdfUnderstandingTableColumn(left + 100D, left + 180D)
            },
            new IReadOnlyList<string>[] {
                new[] { "A" + index, "B" + index },
                new[] { "C" + index, "D" + index }
            },
            Array.Empty<PdfUnderstandingLine>());
    }

    private static PdfUnderstandingTableCandidate CreateCandidate(
        string detectionKind,
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        IReadOnlyList<PdfTextSpan> sourceRuns) =>
        new PdfUnderstandingTableCandidate(
            detectionKind,
            710D,
            650D,
            new[] {
                new PdfUnderstandingTableColumn(10D, 60D),
                new PdfUnderstandingTableColumn(70D, 140D)
            },
            new IReadOnlyList<string>[] {
                new[] { sourceRuns[0].Text, sourceRuns[sourceRuns.Count - 1].Text },
                new[] { "x", "y" }
            },
            sourceLines);

    private static PdfUnderstandingTableCandidate CreateTaggedCandidate(
        string value,
        IReadOnlyList<PdfUnderstandingLine> sourceLines,
        IReadOnlyList<PdfTextSpan> sourceRuns) =>
        PdfUnderstandingTableCandidate.FromTagged(
            710D,
            650D,
            new[] {
                new PdfUnderstandingTableColumn(10D, 60D),
                new PdfUnderstandingTableColumn(70D, 140D)
            },
            new IReadOnlyList<string>[] {
                new[] { value, "Value" },
                new[] { value + "2", "Value2" }
            },
            sourceLines,
            sourceRuns,
            0.99D,
            new[] { new PdfInferenceEvidence("table.tagged-structure", "Tagged test table.", 0.99D) },
            static _ => { },
            static () => { });

    private static PdfUnderstandingLine CreateLine(string text, params PdfTextSpan[] runs) {
        var word = new PdfUnderstandingWord(
            text,
            runs.Min(static run => run.X),
            runs.Max(static run => run.X + run.Advance),
            runs.Average(static run => run.Y),
            runs.Max(static run => run.FontSize),
            0D,
            runs);
        return new PdfUnderstandingLine(new[] { word });
    }

    private sealed class BaselineMergedWordGroupingStage : IPdfWordGroupingStage {
        public IReadOnlyList<PdfUnderstandingWord> GroupWords(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfTextSpan> runs) {
            var words = new List<PdfUnderstandingWord>();
            foreach (IGrouping<double, PdfTextSpan> baseline in runs.GroupBy(static run => run.Y)) {
                PdfTextSpan[] ordered = baseline.OrderBy(static run => run.X).ToArray();
                context.ConsumeWork(ordered.Length + 1L);
                words.Add(new PdfUnderstandingWord(
                    string.Join(" ", ordered.Select(static run => run.Text)),
                    ordered[0].X,
                    ordered.Max(static run => run.X + Math.Max(0D, run.Advance)),
                    baseline.Key,
                    ordered.Max(static run => run.FontSize),
                    0D,
                    Array.AsReadOnly(ordered)));
            }
            return words.AsReadOnly();
        }
    }
}
