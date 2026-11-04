using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;

namespace PSWriteOffice.Foundation;

/// <summary>Runs bounded task work while keeping PowerShell stream writes on the existing pipeline bridge.</summary>
internal static class OfficeBatch {
    internal static async Task RunAsync<T>(IEnumerable<T> inputs, Func<T, CancellationToken, Task<object?>> worker,
        int throttleLimit, CancellationToken cancellationToken, IAsyncCmdletPipeline pipeline, string activity, string diagnosticCode) {
        if (throttleLimit < 1) throw new ArgumentOutOfRangeException(nameof(throttleLimit));
        if (inputs == null) throw new ArgumentNullException(nameof(inputs));
        if (worker == null) throw new ArgumentNullException(nameof(worker));
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        using var iterator = inputs.GetEnumerator();
        var pending = new List<Task<Outcome<T>>>();
        int completed = 0;
        bool exhausted = false;
        try {
            while (pending.Count != 0 || !exhausted) {
                stop.Token.ThrowIfCancellationRequested();
                while (!exhausted && pending.Count < throttleLimit) {
                    if (!iterator.MoveNext()) { exhausted = true; break; }
                    T input = iterator.Current;
                    pending.Add(Task.Run(async () => {
                        try { return new Outcome<T>(input, await worker(input, stop.Token).ConfigureAwait(false), null); }
                        catch (OperationCanceledException) when (stop.IsCancellationRequested) { throw; }
                        catch (Exception error) { return new Outcome<T>(input, null, error); }
                    }, stop.Token));
                }
                if (pending.Count == 0) break;
                Task<Outcome<T>> done = await Task.WhenAny(pending).ConfigureAwait(false);
                pending.Remove(done);
                Outcome<T> result = await done.ConfigureAwait(false);
                stop.Token.ThrowIfCancellationRequested();
                if (result.Error != null) pipeline.WriteError(OfficeError.Create(result.Error, diagnosticCode, result.Input));
                else if (result.Value != null) pipeline.WriteObject(result.Value);
                pipeline.WriteProgress(new ProgressRecord(0, activity, $"{++completed} completed") { PercentComplete = -1 });
            }
            pipeline.WriteProgress(new ProgressRecord(0, activity, $"{completed} completed") { RecordType = ProgressRecordType.Completed });
        } finally {
            stop.Cancel();
            // All producers finish before cancellation resources or the pipeline lifecycle can end.
            try { await Task.WhenAll(pending).ConfigureAwait(false); }
            catch (OperationCanceledException) when (stop.IsCancellationRequested) { }
        }
    }
    private sealed class Outcome<T> {
        internal Outcome(T input, object? value, Exception? error) { Input = input; Value = value; Error = error; }
        internal T Input { get; }
        internal object? Value { get; }
        internal Exception? Error { get; }
    }
}
