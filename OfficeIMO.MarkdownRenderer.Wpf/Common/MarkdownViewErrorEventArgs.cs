namespace OfficeIMO.MarkdownRenderer.Wpf;

/// <summary>
/// Describes an exception raised while the markdown host is initializing, rendering, or handling shell events.
/// </summary>
public sealed class MarkdownViewErrorEventArgs : EventArgs {
    /// <summary>
    /// Creates a new error event args instance.
    /// </summary>
    public MarkdownViewErrorEventArgs(string context, Exception exception) {
        Context = string.IsNullOrWhiteSpace(context)
            ? throw new ArgumentException("Context is required.", nameof(context))
            : context;
        Exception = exception ?? throw new ArgumentNullException(nameof(exception));
    }

    /// <summary>
    /// Short text describing the operation that failed.
    /// </summary>
    public string Context { get; }

    /// <summary>
    /// The exception observed by the markdown host.
    /// </summary>
    public Exception Exception { get; }
}
