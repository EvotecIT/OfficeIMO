using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

/// <summary>Opt-in preview policy for tests that inspect deliberate incomplete fallback artifacts.</summary>
internal static class IWorkTestPolicy {
    internal static IWorkConversionOptions ForIncompletePreview(IWorkConversionOptions options) {
        IWorkConversionOptions preview = options.Clone();
        preview.RequireCompleteVisualCoverage = false;
        return preview;
    }
}
