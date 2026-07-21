using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInteroperabilityCorpusTests {
    [Fact]
    public void CuratedCorpus_ProducesStableMetadataMutationDecisions() {
        IReadOnlyList<PdfInteroperabilityCorpusCase> corpus = PdfInteroperabilityCorpusSupport.Build();

        Assert.Equal(PdfInteroperabilityCorpusSupport.CaseCount, corpus.Count);
        Assert.Equal(corpus.Count, corpus.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.All(corpus, item => {
            Assert.NotEmpty(item.Pdf);
            Assert.NotEmpty(item.Features);
            PdfMutationPlan plan = PdfMutationPlanner.Plan(
                item.Pdf,
                PdfMutationOperation.UpdateMetadata,
                item.ReadOptions);
            Assert.Equal(item.ExpectedMetadataMode, plan.ExecutionMode);
            Assert.NotEmpty(plan.CapabilityRecords);
            Assert.Contains(plan.CapabilityRecords, record => record.Kind == PdfMutationCapabilityKind.MetadataChanges);
        });
    }

    [Fact]
    public void CuratedCorpus_RunsMetadataChangesThroughPreservationMatrix() {
        IReadOnlyList<PdfRewritePreservationMatrixScenario> scenarios = PdfInteroperabilityCorpusSupport.Build()
            .Select(BuildMetadataScenario)
            .ToArray();

        PdfRewritePreservationMatrixReport report = PdfRewritePreservationMatrix.AssertExpected(scenarios);

        Assert.True(report.Passed);
        Assert.Equal(PdfInteroperabilityCorpusSupport.CaseCount, report.Entries.Count);
        Assert.All(report.Entries, entry => {
            Assert.True(entry.Passed, entry.Summary);
            if (entry.ActualClassification == PdfRewritePreservationMatrixClassification.Blocked) {
                Assert.NotNull(entry.FailureMessage);
            } else {
                Assert.NotNull(entry.PreservationReport);
                Assert.True(entry.PreservationReport!.IsPreserved, entry.PreservationReport.Summary);
            }
        });
    }

    private static PdfRewritePreservationMatrixScenario BuildMetadataScenario(PdfInteroperabilityCorpusCase item) {
        return new PdfRewritePreservationMatrixScenario(
                item.Id,
                "PlannedMetadataUpdate",
                item.Pdf,
                pdf => ApplyPlannedMetadataUpdate(pdf, item.ReadOptions)) {
                ExpectedClassification = item.ExpectedMetadataMode == PdfMutationExecutionMode.Blocked
                    ? PdfRewritePreservationMatrixClassification.Blocked
                    : PdfRewritePreservationMatrixClassification.RewriteSafe,
                PreservationOptions = new PdfRewritePreservationOptions {
                    OriginalReadOptions = item.ReadOptions,
                    PreserveSecurityState = !item.Features.Contains("encrypted", StringComparer.Ordinal)
                }.AllowMetadataChanges("Title")
            }
            .WithSourceFeatures(item.Features.ToArray());
    }

    private static byte[] ApplyPlannedMetadataUpdate(byte[] pdf, PdfReadOptions? readOptions) {
        PdfMutationPlan plan = PdfMutationPlanner.Plan(pdf, PdfMutationOperation.UpdateMetadata, readOptions);
        switch (plan.ExecutionMode) {
            case PdfMutationExecutionMode.FullRewrite:
                return PdfMetadataEditor.UpdateMetadata(pdf, "Corpus metadata update", author: null, subject: null, keywords: null, readOptions: readOptions);
            case PdfMutationExecutionMode.AppendOnly:
                return PdfIncrementalUpdater.UpdateMetadata(pdf, title: "Corpus metadata update", readOptions: readOptions);
            default:
                throw new NotSupportedException(plan.Summary);
        }
    }
}
