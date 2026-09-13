using AngleSharp.Dom;
using AngleSharp.Io;

namespace OfficeIMO.Html.Runtime.Worker;

// Owns request cancellation until the response task has settled. Cancel and
// completion may race when a link changes href or its document is disposed.
internal sealed class RuntimeModuleDownload : IDownload {
    private readonly object _sync = new();
    private CancellationTokenSource? _cancellation;

    internal RuntimeModuleDownload(Url target, object source, Task<IResponse> task, CancellationTokenSource cancellation) {
        Target = target;
        Source = source;
        Task = task;
        _cancellation = cancellation;
        _ = task.ContinueWith(static (_, state) => ((RuntimeModuleDownload)state!).Complete(), this,
            CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
    }

    public Url Target { get; }
    public object Source { get; }
    public Task<IResponse> Task { get; }
    public bool IsCompleted => Task.IsCompleted;
    public bool IsRunning => !Task.IsCompleted;

    public void Cancel() {
        lock (_sync) _cancellation?.Cancel();
    }

    private void Complete() {
        lock (_sync) {
            _cancellation?.Dispose();
            _cancellation = null;
        }
    }
}
