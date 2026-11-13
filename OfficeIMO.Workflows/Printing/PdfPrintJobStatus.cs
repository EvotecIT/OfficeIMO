namespace OfficeIMO.Workflows;

/// <summary>State of a submitted print job, reported the same way by every transport.</summary>
public enum PdfPrintJobState {
    /// <summary>Waiting to be processed.</summary>
    Pending,
    /// <summary>Held or paused, and not yet eligible for processing.</summary>
    Held,
    /// <summary>Being processed, spooled or printed.</summary>
    Processing,
    /// <summary>Processing stopped, usually because the printer needs attention. See the state reasons.</summary>
    Stopped,
    /// <summary>Canceled by a user or operator. Terminal.</summary>
    Canceled,
    /// <summary>Aborted by the printer or spooler. Terminal.</summary>
    Aborted,
    /// <summary>Finished; the printer or spooler considers the job delivered. Terminal.</summary>
    Completed,
    /// <summary>The queue no longer reports the job: it finished and was purged, was deleted, or never existed. Terminal.</summary>
    Removed
}

/// <summary>A job's state as reported by its queue or printer.</summary>
/// <param name="PrinterName">The printer name or URI the job was submitted to.</param>
/// <param name="JobId">The receipt's job identifier.</param>
/// <param name="State">Current state.</param>
public sealed record PdfPrintJobStatus(string PrinterName, string JobId, PdfPrintJobState State) {
    /// <summary>Machine-readable reasons, such as <c>job-printing</c>, <c>media-empty</c> or <c>paused</c>.</summary>
    public IReadOnlyList<string> StateReasons { get; init; } = [];
    /// <summary>Human-readable status text from the printer or spooler.</summary>
    public string? StateMessage { get; init; }
    /// <summary>True for canceled, aborted, completed and removed jobs, whose state no longer changes.</summary>
    public bool IsTerminal => State is PdfPrintJobState.Canceled or PdfPrintJobState.Aborted or PdfPrintJobState.Completed or PdfPrintJobState.Removed;
}

/// <summary>Operations composed from the <see cref="IPdfPrinterService"/> members.</summary>
public static partial class PdfPrinterServiceExtensions {
    /// <summary>Polls a submitted job until it reaches a terminal state.</summary>
    /// <param name="printers">Printer service the job was submitted through.</param>
    /// <param name="printerName">The receipt's <see cref="PdfPrintSubmission.PrinterName"/>.</param>
    /// <param name="jobId">The receipt's <see cref="PdfPrintSubmission.JobId"/>.</param>
    /// <param name="timeout">Longest wait, or <see cref="Timeout.InfiniteTimeSpan"/>.</param>
    /// <param name="pollInterval">Delay between queries; defaults to one second.</param>
    /// <param name="cancellationToken">Stops waiting; the job is not canceled.</param>
    /// <returns>The job in its terminal state.</returns>
    /// <exception cref="TimeoutException">The job was not terminal when the timeout elapsed.</exception>
    public static async Task<PdfPrintJobStatus> WaitForJobAsync(this IPdfPrinterService printers, string printerName, string jobId, TimeSpan timeout,
        TimeSpan? pollInterval = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(printers);
        TimeSpan interval = PrintJobPolling.Interval(timeout, pollInterval);
        return await PrintJobPolling.WaitAsync(token => printers.GetJobAsync(printerName, jobId, token), status => status.IsTerminal,
            status => status.State.ToString(), jobId, timeout, interval, cancellationToken).ConfigureAwait(false);
    }
}

/// <summary>Polls a job query until it reports a terminal state, bounding every query by the remaining wait.</summary>
internal static class PrintJobPolling {
    /// <summary>Validates the wait arguments and returns the poll interval (one second by default).</summary>
    public static TimeSpan Interval(TimeSpan timeout, TimeSpan? pollInterval) {
        if (timeout < TimeSpan.Zero && timeout != Timeout.InfiniteTimeSpan) throw new ArgumentOutOfRangeException(nameof(timeout));
        TimeSpan interval = pollInterval ?? TimeSpan.FromSeconds(1);
        if (interval <= TimeSpan.Zero || interval > TimeSpan.FromHours(1)) throw new ArgumentOutOfRangeException(nameof(pollInterval));
        return interval;
    }

    /// <summary>
    /// Queries until <paramref name="isTerminal"/> holds. Each query receives a token that expires at the deadline, and the wait
    /// stops at the deadline even when a query ignores its token; a query that expires, or a terminal answer that arrives after
    /// the deadline, raises <see cref="TimeoutException"/>. Caller cancellation raises <see cref="OperationCanceledException"/>.
    /// </summary>
    public static async Task<T> WaitAsync<T>(Func<CancellationToken, Task<T>> query, Func<T, bool> isTerminal, Func<T, string> describe,
        string jobId, TimeSpan timeout, TimeSpan interval, CancellationToken cancellationToken) {
        bool bounded = timeout != Timeout.InfiniteTimeSpan;
        var clock = System.Diagnostics.Stopwatch.StartNew();
        string? lastState = null;
        while (true) {
            T result;
            if (!bounded) {
                result = await query(cancellationToken).ConfigureAwait(false);
            } else {
                TimeSpan remaining = timeout - clock.Elapsed;
                if (remaining <= TimeSpan.Zero) throw Expired(jobId, timeout, lastState);
                using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                deadline.CancelAfter(remaining);
                Task<T> pending = query(deadline.Token);
                if (!pending.IsCompleted) {
                    using var stopTimer = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    Task timer = Task.Delay(remaining, stopTimer.Token);
                    if (await Task.WhenAny(pending, timer).ConfigureAwait(false) != pending) {
                        // The query ignored its token. Stop waiting for it, but observe its eventual failure.
                        _ = pending.ContinueWith(static task => _ = task.Exception, CancellationToken.None,
                            TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
                        cancellationToken.ThrowIfCancellationRequested();
                        throw Expired(jobId, timeout, lastState);
                    }
                    stopTimer.Cancel();
                }
                try {
                    result = await pending.ConfigureAwait(false);
                } catch (OperationCanceledException error) when (!cancellationToken.IsCancellationRequested && deadline.IsCancellationRequested) {
                    throw Expired(jobId, timeout, lastState, error);
                }
                if (clock.Elapsed > timeout) throw Expired(jobId, timeout, describe(result));
            }
            if (isTerminal(result)) return result;
            lastState = describe(result);
            TimeSpan delay = interval;
            if (bounded) {
                TimeSpan remaining = timeout - clock.Elapsed;
                if (remaining <= TimeSpan.Zero) throw Expired(jobId, timeout, lastState);
                if (remaining < delay) delay = remaining;
            }
            await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
        }
    }

    private static TimeoutException Expired(string jobId, TimeSpan timeout, string? state, Exception? inner = null) =>
        new(state is null ? $"Job {jobId} did not report a terminal state within {timeout}." : $"Job {jobId} was still {state} after {timeout}.", inner);
}
