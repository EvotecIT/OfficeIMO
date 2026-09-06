using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private readonly object _capabilityGate = new();
    private readonly Dictionary<PdfMutationOperation, bool> _capabilities = [];
    private byte[]? _capabilitySource;

    private bool CanPlan(PdfMutationOperation operation) {
        lock (_capabilityGate) {
            if (_disposed) return false;
            byte[] source = _bytes;
            // Workspace bytes are immutable snapshots. Edits, recovery, undo, and redo replace
            // the array, so command queries can share a result only for that exact snapshot.
            if (!ReferenceEquals(_capabilitySource, source)) {
                _capabilities.Clear();
                _capabilitySource = source;
            }
            if (_capabilities.TryGetValue(operation, out bool allowed)) return allowed;
            try { allowed = LoadDocument(source).PlanMutation(operation).CanExecute; }
            catch { allowed = false; }
            _capabilities[operation] = allowed;
            return allowed;
        }
    }

    private void ClearCapabilityCache() {
        lock (_capabilityGate) {
            _capabilities.Clear();
            _capabilitySource = null;
        }
    }
}
