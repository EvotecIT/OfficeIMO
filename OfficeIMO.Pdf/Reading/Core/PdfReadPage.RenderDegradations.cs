using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal sealed partial class PageContentBudget {
        private const int MaximumRenderDegradations = 64;
        private const int MaximumRenderDegradationSubjectLength = 128;
        private List<PdfRenderCapabilityDiagnostic>? _renderDegradations;
        private HashSet<string>? _renderDegradationKeys;

        /// <summary>Content skipped while drawing the page, which the capability pre-scan cannot predict.</summary>
        internal IReadOnlyList<PdfRenderCapabilityDiagnostic> RenderDegradations =>
            _renderDegradations ?? (IReadOnlyList<PdfRenderCapabilityDiagnostic>)Array.Empty<PdfRenderCapabilityDiagnostic>();

        internal void RecordRenderDegradation(string capabilityId, string? subject) {
            if (subject is { Length: > MaximumRenderDegradationSubjectLength }) {
                subject = subject.Substring(0, MaximumRenderDegradationSubjectLength);
            }
            _renderDegradationKeys ??= new HashSet<string>(StringComparer.Ordinal);
            if (_renderDegradationKeys.Count >= MaximumRenderDegradations ||
                !_renderDegradationKeys.Add(capabilityId + "\u0000" + subject)) return;
            (_renderDegradations ??= new List<PdfRenderCapabilityDiagnostic>())
                .Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(capabilityId), subject));
        }
    }

    /// <summary>
    /// Whether the current thread can enter one more level of nested content. Each nesting level of the content walkers
    /// uses tens of kilobytes of stack, so real documents with a few dozen nested transparency groups would otherwise
    /// exhaust a 1 MiB thread stack long before <see cref="PdfReadLimits.MaxContentNestingDepth"/> applies.
    /// </summary>
    private static bool HasSufficientExecutionStack() {
#if NETSTANDARD2_0 || NET472
        try {
            RuntimeHelpers.EnsureSufficientExecutionStack();
            return true;
        } catch (InsufficientExecutionStackException) {
            return false;
        }
#else
        return RuntimeHelpers.TryEnsureSufficientExecutionStack();
#endif
    }
}
