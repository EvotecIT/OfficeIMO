namespace OfficeIMO.Pdf;

/// <summary>
/// Runs reusable rewrite-preservation proof matrices for PDF manipulation operations.
/// </summary>
public static class PdfRewritePreservationMatrix {
    /// <summary>
    /// Runs the supplied scenarios and classifies each rewrite outcome.
    /// </summary>
    public static PdfRewritePreservationMatrixReport Assess(IEnumerable<PdfRewritePreservationMatrixScenario> scenarios) {
        Guard.NotNull(scenarios, nameof(scenarios));

        var entries = new List<PdfRewritePreservationMatrixEntry>();
        foreach (PdfRewritePreservationMatrixScenario scenario in scenarios) {
            Guard.NotNull(scenario, nameof(scenarios));
            entries.Add(AssessScenario(scenario));
        }

        return new PdfRewritePreservationMatrixReport(entries.AsReadOnly());
    }

    /// <summary>
    /// Runs the supplied scenarios and throws when any scenario produces an unexpected classification.
    /// </summary>
    public static PdfRewritePreservationMatrixReport AssertExpected(IEnumerable<PdfRewritePreservationMatrixScenario> scenarios) {
        PdfRewritePreservationMatrixReport report = Assess(scenarios);
        report.ThrowIfFailed();
        return report;
    }

    private static PdfRewritePreservationMatrixEntry AssessScenario(PdfRewritePreservationMatrixScenario scenario) {
        IReadOnlyList<string> sourceFeatures = scenario.SourceFeatures.ToArray();
        try {
            byte[] rewritten = scenario.Rewrite((byte[])scenario.SourcePdf.Clone());
            PdfRewritePreservationReport preservationReport = PdfRewritePreservation.Assess(scenario.SourcePdf, rewritten, scenario.PreservationOptions);
            PdfRewritePreservationMatrixClassification classification = preservationReport.IsPreserved
                ? PdfRewritePreservationMatrixClassification.RewriteSafe
                : PdfRewritePreservationMatrixClassification.PreservationFailed;

            return new PdfRewritePreservationMatrixEntry(
                scenario.Id,
                scenario.Operation,
                scenario.ExpectedClassification,
                classification,
                sourceFeatures,
                preservationReport,
                failureType: null,
                failureMessage: null);
        } catch (PdfEncryptionException exception) {
            return BuildFailureEntry(scenario, sourceFeatures, PdfRewritePreservationMatrixClassification.Blocked, exception);
        } catch (NotSupportedException exception) {
            return BuildFailureEntry(scenario, sourceFeatures, PdfRewritePreservationMatrixClassification.Blocked, exception);
        } catch (ArgumentException exception) {
            return BuildFailureEntry(scenario, sourceFeatures, PdfRewritePreservationMatrixClassification.OperationFailed, exception);
        } catch (InvalidOperationException exception) {
            return BuildFailureEntry(scenario, sourceFeatures, PdfRewritePreservationMatrixClassification.OperationFailed, exception);
        }
    }

    private static PdfRewritePreservationMatrixEntry BuildFailureEntry(
        PdfRewritePreservationMatrixScenario scenario,
        IReadOnlyList<string> sourceFeatures,
        PdfRewritePreservationMatrixClassification classification,
        Exception exception) {
        return new PdfRewritePreservationMatrixEntry(
            scenario.Id,
            scenario.Operation,
            scenario.ExpectedClassification,
            classification,
            sourceFeatures,
            preservationReport: null,
            failureType: exception.GetType().Name,
            failureMessage: exception.Message);
    }
}
