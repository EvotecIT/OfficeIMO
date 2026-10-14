using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

/// <summary>Explicit policies and assertions for tests that inspect qualified conversion limits.</summary>
internal static class IWorkTestPolicy {
    internal static void AssertCellPaddingOmission(IWorkConversionReport report) {
        Assert.True(report.IsPartialEditableReconstruction);
        IWorkDiagnostic omission = Assert.Single(report.Diagnostics,
            diagnostic => diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission
                && diagnostic.Code != "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION");
        Assert.Equal("IWORK_NUMBERS_CELL_PADDING_OMITTED", omission.Code);
        Assert.DoesNotContain(report.Diagnostics,
            diagnostic => diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed
                && diagnostic.Code != "IWORK_RECORD_FIDELITY_UNASSESSED");
        Assert.Throws<InvalidOperationException>(() => report.RequireCompleteEditableReconstruction());
    }

    internal static IWorkConversionOptions ForIncompletePreview(IWorkConversionOptions options) {
        IWorkConversionOptions preview = options.Clone();
        preview.RequireCompleteVisualCoverage = false;
        return preview;
    }
}
