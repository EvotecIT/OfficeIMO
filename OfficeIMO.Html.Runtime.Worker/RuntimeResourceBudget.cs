namespace OfficeIMO.Html.Runtime.Worker;

// Authority and accounting belong to the browsing session. Each document gets a
// cancellable loader, but replacing that document never resets these limits.
internal sealed class RuntimeResourceBudget(HtmlScriptRequest options) : IDisposable {
    internal readonly object Sync = new();
    internal readonly Dictionary<string, HtmlRuntimeResource> Loaded = new(StringComparer.Ordinal);
    internal readonly HashSet<string> Origins = new(options.ResourcePolicy.AllowedOrigins.Select(HtmlRuntimeResourcePolicy.Origin)
        .Append(HtmlRuntimeResourcePolicy.Origin(options.DocumentUrl)), StringComparer.OrdinalIgnoreCase);
    internal readonly SemaphoreSlim Concurrency = new(options.ResourcePolicy.MaxConcurrentRequests);
    internal long ReceivedBytes;
    internal long SentBytes;
    internal long Requests;
    private int _activeOperations;
    private bool _stopping;
    private TaskCompletionSource _idle = Completed();

    internal void BeginOperation() {
        lock (Sync) {
            if (_stopping) throw new OperationCanceledException("The browsing session is stopping.");
            if (_activeOperations++ == 0) _idle = new(TaskCreationOptions.RunContinuationsAsynchronously);
        }
    }

    internal void EndOperation() {
        lock (Sync) {
            if (--_activeOperations == 0) _idle.TrySetResult();
        }
    }

    internal Task StopAndWaitAsync() { lock (Sync) { _stopping = true; return _idle.Task; } }

    public void Dispose() => Concurrency.Dispose();

    private static TaskCompletionSource Completed() {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        completion.SetResult();
        return completion;
    }
}
