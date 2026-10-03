using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Operation-local budgets and nested result capture. Iterator callers restore this scope on every move.</summary>
internal sealed class ReaderReadScope : IDisposable {
    private static readonly AsyncLocal<ReaderReadScope?> Active = new AsyncLocal<ReaderReadScope?>();
    private readonly ReaderReadScope? _previous;
    private readonly List<OfficeDocumentNestedResult> _nested;
    internal static ReaderReadScope? Current => Active.Value;
    internal ReaderOperationBudget? Budget { get; }
    internal int Depth { get; }
    private ReaderReadScope(ReaderOperationBudget? budget, int depth, bool activate,
        List<OfficeDocumentNestedResult>? nestedResults = null) {
        Budget = budget; Depth = depth;
        _nested = nestedResults ?? new List<OfficeDocumentNestedResult>();
        if (activate) { _previous = Active.Value; Active.Value = this; }
    }
    internal static ReaderReadScope Enter(ReaderOptions? options, bool nested = false) {
        var previous = Current;
        var budget = previous?.Budget ?? (options?.ResourceLimits == null ? null : new ReaderOperationBudget(options.ResourceLimits));
        int depth = (previous?.Depth ?? 0) + (nested ? 1 : 0);
        budget?.CheckDepth(depth);
        return new ReaderReadScope(budget, depth, activate: true);
    }
    internal static ReaderReadScope CreateDetached(ReaderOptions? options) => new ReaderReadScope(
        Current?.Budget ?? (options?.ResourceLimits == null ? null : new ReaderOperationBudget(options.ResourceLimits)),
        Current?.Depth ?? 0, activate: false);
    internal static ReaderReadScope EnterContainer(ReaderOptions options) {
        var previous = Current;
        var budget = previous?.Budget ?? (options.ResourceLimits == null ? null : new ReaderOperationBudget(options.ResourceLimits));
        int depth = (previous?.Depth ?? 0) + 1;
        budget?.CheckDepth(depth);
        return new ReaderReadScope(budget, depth, activate: true, nestedResults: previous?._nested);
    }
    internal static IDisposable Use(ReaderReadScope scope) {
        var previous = Active.Value; Active.Value = scope; return new Restore(previous);
    }
    internal static OfficeDocumentReadResult Complete(OfficeDocumentReadResult result) {
        var scope = Current;
        if (scope == null) return result;
        lock (scope._nested) if (scope._nested.Count > 0)
            result.NestedDocuments = result.NestedDocuments.Concat(scope._nested).Distinct().ToArray();
        scope.Budget?.AddDocument(result);
        return result;
    }
    internal static void AttachPendingNested(OfficeDocumentReadResult result) {
        var scope = Current;
        if (scope == null) return;
        lock (scope._nested) {
            if (scope._nested.Count == 0) return;
            result.NestedDocuments = result.NestedDocuments.Concat(scope._nested).Distinct().ToArray();
            scope._nested.Clear();
        }
    }
    internal static void RecordNested(string path, OfficeDocumentReadResult document) {
        var scope = Current;
        if (scope == null) return;
        lock (scope._nested) scope._nested.Add(new OfficeDocumentNestedResult { Path = path, Document = document });
    }
    public void Dispose() { Active.Value = _previous; }
    private sealed class Restore : IDisposable {
        private readonly ReaderReadScope? _previous;
        internal Restore(ReaderReadScope? previous) { _previous = previous; }
        public void Dispose() { Active.Value = _previous; }
    }
}
