using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfMutationPortfolioCancellationTests {
    [Fact]
    public void CanceledAssessmentDoesNotEnumerateCallerRequests() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var document = PdfDocument.Create();

        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            document.AssessMutations(cancellation.Token, UnexpectedOperations(), UnexpectedFields()));

        Assert.Equal(cancellation.Token, error.CancellationToken);

        static IEnumerable<PdfMutationOperation> UnexpectedOperations() {
            yield return Unexpected<PdfMutationOperation>();
        }
        static IEnumerable<string> UnexpectedFields() {
            yield return Unexpected<string>();
        }
        static T Unexpected<T>() => throw new InvalidOperationException("Canceled assessment must not enumerate requests.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AssessmentStopsWhenCallerCancelsDuringRequestEnumeration(bool cancelDuringFields) {
        using var cancellation = new CancellationTokenSource();
        var document = PdfDocument.Create();

        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            document.AssessMutations(cancellation.Token, Operations(), Fields()));

        Assert.Equal(cancellation.Token, error.CancellationToken);

        IEnumerable<PdfMutationOperation> Operations() {
            if (!cancelDuringFields) cancellation.Cancel();
            yield return PdfMutationOperation.UpdateMetadata;
            if (!cancelDuringFields) throw new InvalidOperationException("Assessment continued after cancellation.");
        }
        IEnumerable<string> Fields() {
            cancellation.Cancel();
            yield return "Name";
            throw new InvalidOperationException("Assessment continued after cancellation.");
        }
    }

    [Fact]
    public void CancellablePortfolioRetainsOperationOrderBlockersAndSharedPreflight() {
        PdfDocument document = PdfDocument.Load(PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf());
        var operations = new[] {
            PdfMutationOperation.ModifyPageContent,
            PdfMutationOperation.UpdateMetadata,
            PdfMutationOperation.ModifyPageContent
        };
        PdfMutationPortfolioReport baseline = document.AssessMutations(operations);
        PdfMutationPortfolioReport result = document.AssessMutations(CancellationToken.None, operations);

        Assert.Equal(baseline.Plans.Select(plan => plan.Operation), result.Plans.Select(plan => plan.Operation));
        Assert.Equal(2, result.Plans.Count);
        Assert.Contains(result.Plans, plan => plan.CanExecute);
        Assert.Contains(result.Plans, plan => !plan.CanExecute && plan.BlockerCodes.Count > 0);
        foreach (PdfMutationPlan plan in result.Plans) {
            PdfMutationPlan expected = baseline.Get(plan.Operation);
            Assert.Equal(expected.ExecutionMode, plan.ExecutionMode);
            Assert.Equal(expected.BlockerCodes, plan.BlockerCodes);
            Assert.Same(result.Preflight, plan.Preflight);
        }
    }
}
