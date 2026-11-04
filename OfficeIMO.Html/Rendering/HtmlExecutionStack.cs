using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace OfficeIMO.Html;

/// <summary>
/// Supplies the stack that <see cref="HtmlRenderOptions.MaxLayoutDepth"/> permits. A stack overflow
/// ends the process and no caller can catch it, and worker threads can start with as little as 512 KB.
/// </summary>
internal static class HtmlExecutionStack {
    // Each layout level costs up to about 30 KB of unoptimized frames.
    private const int StackBytesPerLayoutLevel = 64 * 1024;
    internal const int MinimumStackBytes = 16 * 1024 * 1024;
    private const int MaximumStackBytes = 1024 * 1024 * 1024;

    internal static int StackBytesFor(int layoutDepth) =>
        (int)Math.Min(MaximumStackBytes, Math.Max(MinimumStackBytes, (long)layoutDepth * StackBytesPerLayoutLevel));

    internal static bool HasRoom() {
#if NET8_0_OR_GREATER
        return RuntimeHelpers.TryEnsureSufficientExecutionStack();
#else
        try {
            RuntimeHelpers.EnsureSufficientExecutionStack();
            return true;
        } catch (InsufficientExecutionStackException) {
            return false;
        }
#endif
    }

    /// <summary>Runs <paramref name="work"/> on a new stack while the caller waits, so execution stays sequential.</summary>
    /// <returns><see langword="false"/> without running the work when the host cannot start a thread.</returns>
    internal static bool TryRun<T>(Func<T> work, int stackBytes, out T result) {
        T completed = default!;
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() => {
            try {
                completed = work();
            } catch (Exception exception) {
                failure = ExceptionDispatchInfo.Capture(exception);
            }
        }, stackBytes) {
            IsBackground = true,
            Name = "OfficeIMO HTML rendering",
            CurrentCulture = CultureInfo.CurrentCulture,
            CurrentUICulture = CultureInfo.CurrentUICulture
        };
        try {
            thread.Start();
        } catch (Exception exception) when (exception is PlatformNotSupportedException or ThreadStartException or OutOfMemoryException) {
            result = default!;
            return false;
        }
        thread.Join();
        failure?.Throw();
        result = completed;
        return true;
    }
}
