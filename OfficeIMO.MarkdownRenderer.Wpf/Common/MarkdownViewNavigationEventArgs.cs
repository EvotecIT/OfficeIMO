namespace OfficeIMO.MarkdownRenderer.Wpf;

/// <summary>
/// Describes a navigation request initiated from inside the markdown WebView.
/// </summary>
public sealed class MarkdownViewNavigationEventArgs : EventArgs {
    /// <summary>
    /// Creates a new navigation request event args instance.
    /// </summary>
    public MarkdownViewNavigationEventArgs(Uri uri) {
        Uri = uri ?? throw new ArgumentNullException(nameof(uri));
    }

    /// <summary>
    /// The absolute URI requested by the embedded markdown surface.
    /// </summary>
    public Uri Uri { get; }

    /// <summary>
    /// Set to <see langword="true"/> when the host handled the navigation itself.
    /// </summary>
    public bool Handled { get; set; }
}
