using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;

namespace OfficeIMO.MarkdownRenderer.Wpf;

/// <summary>
/// WPF/WebView2 markdown host control built on <see cref="OfficeIMO.MarkdownRenderer.MarkdownRenderer"/>.
/// </summary>
public partial class MarkdownView : UserControl, IDisposable {
    /// <summary>
    /// Dependency property backing <see cref="Markdown"/>.
    /// </summary>
    public static readonly DependencyProperty MarkdownProperty =
        DependencyProperty.Register(
            nameof(Markdown),
            typeof(string),
            typeof(MarkdownView),
            new PropertyMetadata(string.Empty, OnMarkdownChanged));

    /// <summary>
    /// Dependency property backing <see cref="BodyHtml"/>.
    /// </summary>
    public static readonly DependencyProperty BodyHtmlProperty =
        DependencyProperty.Register(
            nameof(BodyHtml),
            typeof(string),
            typeof(MarkdownView),
            new PropertyMetadata(string.Empty, OnBodyHtmlChanged));

    /// <summary>
    /// Dependency property backing <see cref="DocumentTitle"/>.
    /// </summary>
    public static readonly DependencyProperty DocumentTitleProperty =
        DependencyProperty.Register(
            nameof(DocumentTitle),
            typeof(string),
            typeof(MarkdownView),
            new PropertyMetadata("Markdown", OnShellPropertyChanged));

    /// <summary>
    /// Dependency property backing <see cref="BaseHref"/>.
    /// </summary>
    public static readonly DependencyProperty BaseHrefProperty =
        DependencyProperty.Register(
            nameof(BaseHref),
            typeof(string),
            typeof(MarkdownView),
            new PropertyMetadata(string.Empty, OnShellPropertyChanged));

    /// <summary>
    /// Dependency property backing <see cref="ShellCss"/>.
    /// </summary>
    public static readonly DependencyProperty ShellCssProperty =
        DependencyProperty.Register(
            nameof(ShellCss),
            typeof(string),
            typeof(MarkdownView),
            new PropertyMetadata(string.Empty, OnShellPropertyChanged));

    /// <summary>
    /// Dependency property backing <see cref="Preset"/>.
    /// </summary>
    public static readonly DependencyProperty PresetProperty =
        DependencyProperty.Register(
            nameof(Preset),
            typeof(MarkdownViewPreset),
            typeof(MarkdownView),
            new PropertyMetadata(MarkdownViewPreset.Strict, OnShellPropertyChanged));

    /// <summary>
    /// Dependency property backing <see cref="OpenLinksExternally"/>.
    /// </summary>
    public static readonly DependencyProperty OpenLinksExternallyProperty =
        DependencyProperty.Register(
            nameof(OpenLinksExternally),
            typeof(bool),
            typeof(MarkdownView),
            new PropertyMetadata(true));

    private readonly SemaphoreSlim _renderGate = new(1, 1);
    private TaskCompletionSource<bool>? _navigationCompletionSource;
    private Action<OfficeIMO.MarkdownRenderer.MarkdownRendererOptions>? _configureRendererOptions;
    private bool _webViewReady;
    private bool _browserEventsAttached;
    private bool _pendingShellReload = true;
    private bool _pendingBodyReload = true;
    private bool _disposed;

    /// <summary>
    /// Creates a new markdown host control instance.
    /// </summary>
    public MarkdownView() {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    /// <summary>
    /// Raised when the embedded surface requests navigation away from the host page.
    /// </summary>
    public event EventHandler<MarkdownViewNavigationEventArgs>? NavigationRequested;

    /// <summary>
    /// Raised when an asynchronous render or shell operation fails.
    /// </summary>
    public event EventHandler<MarkdownViewErrorEventArgs>? ErrorOccurred;

    /// <summary>
    /// Markdown text rendered into the embedded WebView shell.
    /// </summary>
    public string Markdown {
        get => (string)GetValue(MarkdownProperty);
        set => SetValue(MarkdownProperty, value);
    }

    /// <summary>
    /// Optional pre-rendered HTML body inserted into the markdown shell instead of rendering <see cref="Markdown"/>.
    /// This is intended for advanced hosts that compose HTML from multiple rendered markdown fragments.
    /// </summary>
    public string BodyHtml {
        get => (string)GetValue(BodyHtmlProperty);
        set => SetValue(BodyHtmlProperty, value);
    }

    /// <summary>
    /// Title used when building the HTML shell document.
    /// </summary>
    public string DocumentTitle {
        get => (string)GetValue(DocumentTitleProperty);
        set => SetValue(DocumentTitleProperty, value);
    }

    /// <summary>
    /// Optional base href applied to rendered links and images.
    /// </summary>
    public string BaseHref {
        get => (string)GetValue(BaseHrefProperty);
        set => SetValue(BaseHrefProperty, value);
    }

    /// <summary>
    /// Optional host-provided CSS appended to the shell after built-in renderer styles.
    /// </summary>
    public string ShellCss {
        get => (string)GetValue(ShellCssProperty);
        set => SetValue(ShellCssProperty, value);
    }

    /// <summary>
    /// Built-in renderer preset used as the control's baseline configuration.
    /// </summary>
    public MarkdownViewPreset Preset {
        get => (MarkdownViewPreset)GetValue(PresetProperty);
        set => SetValue(PresetProperty, value);
    }

    /// <summary>
    /// When true, external link clicks launch through the operating system shell when not handled by the host.
    /// </summary>
    public bool OpenLinksExternally {
        get => (bool)GetValue(OpenLinksExternallyProperty);
        set => SetValue(OpenLinksExternallyProperty, value);
    }

    /// <summary>
    /// Optional callback invoked with the effective renderer options before the shell or body is rendered.
    /// </summary>
    public Action<OfficeIMO.MarkdownRenderer.MarkdownRendererOptions>? ConfigureRendererOptions {
        get => _configureRendererOptions;
        set {
            ThrowIfDisposed();
            _configureRendererOptions = value;
            QueueRender(rebuildShell: true);
        }
    }

    /// <summary>
    /// Refreshes the currently loaded markdown body inside the existing shell.
    /// </summary>
    public Task RefreshAsync() {
        ThrowIfDisposed();
        _pendingBodyReload = true;
        return RenderPendingAsync();
    }

    /// <summary>
    /// Rebuilds the HTML shell and then refreshes the rendered markdown body.
    /// </summary>
    public Task RebuildShellAsync() {
        ThrowIfDisposed();
        _pendingShellReload = true;
        _pendingBodyReload = true;
        return RenderPendingAsync();
    }

    /// <summary>
    /// Releases unmanaged resources used by the embedded WebView host.
    /// </summary>
    public void Dispose() {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Releases managed resources used by the embedded WebView host.
    /// </summary>
    protected virtual void Dispose(bool disposing) {
        if (_disposed) {
            return;
        }

        if (disposing) {
            Loaded -= OnLoaded;
            DetachBrowserEvents();
            _navigationCompletionSource?.TrySetCanceled();
            _navigationCompletionSource = null;
            _renderGate.Dispose();
            Browser.Dispose();
        }

        _disposed = true;
    }

    private void OnLoaded(object sender, RoutedEventArgs e) {
        QueueRender(rebuildShell: true);
    }

    private static void OnMarkdownChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e) {
        if (dependencyObject is MarkdownView view) {
            view.QueueRender(rebuildShell: false);
        }
    }

    private static void OnBodyHtmlChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e) {
        if (dependencyObject is MarkdownView view) {
            view.QueueRender(rebuildShell: false);
        }
    }

    private static void OnShellPropertyChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e) {
        if (dependencyObject is MarkdownView view) {
            view.QueueRender(rebuildShell: true);
        }
    }

    private void QueueRender(bool rebuildShell) {
        if (_disposed) {
            return;
        }

        if (rebuildShell) {
            _pendingShellReload = true;
        }

        _pendingBodyReload = true;
        if (IsLoaded) {
            _ = RenderPendingSafeAsync();
        }
    }

    private async Task RenderPendingSafeAsync() {
        try {
            await RenderPendingAsync().ConfigureAwait(true);
        } catch (Exception exception) {
            ReportError("render markdown preview", exception, showOverlay: true);
        }
    }

    private async Task RenderPendingAsync() {
        ThrowIfDisposed();
        if (!IsLoaded) {
            return;
        }

        await _renderGate.WaitAsync().ConfigureAwait(true);
        try {
            while (_pendingShellReload || _pendingBodyReload) {
                ThrowIfDisposed();
                var rebuildShell = _pendingShellReload || !_webViewReady;
                _pendingShellReload = false;
                _pendingBodyReload = false;

                ShowStatus("Loading markdown preview...");
                await EnsureWebViewAsync().ConfigureAwait(true);

                var options = CreateEffectiveOptions();
                if (rebuildShell) {
                    await NavigateShellAsync(options).ConfigureAwait(true);
                }

                await UpdateBodyAsync(options).ConfigureAwait(true);
                ShowBrowser();
            }
        } finally {
            _renderGate.Release();
        }
    }

    private async Task EnsureWebViewAsync() {
        if (_webViewReady) {
            return;
        }

        await Browser.EnsureCoreWebView2Async().ConfigureAwait(true);

        var settings = Browser.CoreWebView2.Settings;
        settings.IsStatusBarEnabled = false;
        settings.AreDefaultContextMenusEnabled = false;
        settings.AreDevToolsEnabled = false;

        AttachBrowserEvents();
        _webViewReady = true;
    }

    private async Task NavigateShellAsync(OfficeIMO.MarkdownRenderer.MarkdownRendererOptions options) {
        if (Browser.CoreWebView2 is null) {
            throw new InvalidOperationException("WebView2 is not initialized.");
        }

        _navigationCompletionSource = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        Browser.NavigateToString(OfficeIMO.MarkdownRenderer.MarkdownRenderer.BuildShellHtml(DocumentTitle, options));
        await _navigationCompletionSource.Task.ConfigureAwait(true);
    }

    private Task UpdateBodyAsync(OfficeIMO.MarkdownRenderer.MarkdownRendererOptions options) {
        if (Browser.CoreWebView2 is null) {
            return Task.CompletedTask;
        }

        var bodyHtml = !string.IsNullOrWhiteSpace(BodyHtml)
            ? BodyHtml
            : OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(Markdown ?? string.Empty, options);
        Browser.CoreWebView2.PostWebMessageAsString(bodyHtml);
        return Task.CompletedTask;
    }

    private OfficeIMO.MarkdownRenderer.MarkdownRendererOptions CreateEffectiveOptions() {
        if (!string.IsNullOrWhiteSpace(BaseHref)
            && !MarkdownViewSupport.TryNormalizeBaseHref(BaseHref, out _)) {
            Trace.TraceWarning($"[OfficeIMO.MarkdownRenderer.Wpf] Ignoring invalid BaseHref '{BaseHref}'.");
        }

        return MarkdownViewSupport.CreateEffectiveOptions(Preset, BaseHref, ShellCss, _configureRendererOptions);
    }

    private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e) {
        if (_navigationCompletionSource is null) {
            return;
        }

        if (e.IsSuccess) {
            _navigationCompletionSource.TrySetResult(true);
        } else {
            _navigationCompletionSource.TrySetException(
                new InvalidOperationException($"Markdown shell navigation failed: {e.WebErrorStatus}."));
        }
    }

    private void OnNavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e) {
        if (!e.IsUserInitiated || !MarkdownViewSupport.TryGetExternalNavigationUri(e.Uri, out var navigationUri)) {
            return;
        }

        var args = new MarkdownViewNavigationEventArgs(navigationUri);
        NavigationRequested?.Invoke(this, args);
        e.Cancel = true;

        if (!args.Handled && OpenLinksExternally) {
            TryOpenExternal(navigationUri);
        }
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e) {
        if (MarkdownViewSupport.TryGetClipboardText(e.WebMessageAsJson, out var text)) {
            Clipboard.SetText(text);
        }
    }

    private void AttachBrowserEvents() {
        if (_browserEventsAttached || Browser.CoreWebView2 is null) {
            return;
        }

        Browser.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
        Browser.CoreWebView2.NavigationStarting += OnNavigationStarting;
        Browser.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
        _browserEventsAttached = true;
    }

    private void DetachBrowserEvents() {
        if (!_browserEventsAttached || Browser.CoreWebView2 is null) {
            return;
        }

        Browser.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
        Browser.CoreWebView2.NavigationStarting -= OnNavigationStarting;
        Browser.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
        _browserEventsAttached = false;
    }

    private void TryOpenExternal(Uri navigationUri) {
        try {
            Process.Start(new ProcessStartInfo(navigationUri.AbsoluteUri) {
                UseShellExecute = true
            });
        } catch (Exception exception) {
            Trace.TraceWarning($"[OfficeIMO.MarkdownRenderer.Wpf] Failed to open external URI '{navigationUri}'. {exception}");
        }
    }

    private void ReportError(string context, Exception exception, bool showOverlay) {
        Trace.TraceError($"[OfficeIMO.MarkdownRenderer.Wpf] Failed to {context}. {exception}");
        ErrorOccurred?.Invoke(this, new MarkdownViewErrorEventArgs(context, exception));
        if (showOverlay) {
            ShowError(exception.Message);
        }
    }

    private void ThrowIfDisposed() {
        if (_disposed) {
            throw new ObjectDisposedException(nameof(MarkdownView));
        }
    }

    private void ShowStatus(string text) {
        StatusText.Text = text;
        StatusOverlay.Visibility = Visibility.Visible;
        Browser.Visibility = Visibility.Collapsed;
    }

    private void ShowBrowser() {
        StatusOverlay.Visibility = Visibility.Collapsed;
        Browser.Visibility = Visibility.Visible;
    }

    private void ShowError(string message) {
        StatusText.Text = $"Markdown preview unavailable: {message}";
        StatusOverlay.Visibility = Visibility.Visible;
        Browser.Visibility = Visibility.Collapsed;
    }
}
