namespace OfficeIMO.Html.Runtime;

internal sealed class HtmlProcessRuntimeContext : IHtmlRuntimeContext {
    private readonly HtmlProcessRuntimeProvider _provider;
    private readonly HtmlRuntimeContextOptions _options;
    private readonly object _sync = new();
    private HtmlProcessRuntimeSession? _page;
    private int _disposed;

    internal HtmlProcessRuntimeContext(HtmlProcessRuntimeProvider provider, HtmlRuntimeContextOptions options) {
        _provider = provider;
        _options = options;
        Id = options.Id ?? Guid.NewGuid().ToString("N");
    }

    public string Id { get; }
    public HtmlRuntimeProviderDescriptor Provider => _provider.Descriptor;
    public IReadOnlyList<IHtmlRuntimePage> Pages {
        get {
            lock (_sync) return _page == null ? Array.Empty<IHtmlRuntimePage>() : new IHtmlRuntimePage[] { _page };
        }
    }

    public async Task<IHtmlRuntimePage> OpenPageAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
        HtmlScriptRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        HtmlProcessRuntimeSession page;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
            if (_page != null) throw new InvalidOperationException("This provider supports one page per context.");
            string pageId = Guid.NewGuid().ToString("N");
            page = new HtmlProcessRuntimeSession(_provider.WorkerPath, _provider.DotnetExecutable, _provider.DomServices,
                input, Id, pageId, Provider, _options.Trace, () => PageDisposed(pageId));
            _page = page;
        }
        try {
            await page.OpenAsync(cancellationToken).ConfigureAwait(false);
            return page;
        } catch {
            await page.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    private void PageDisposed(string pageId) {
        lock (_sync) if (_page?.Id == pageId) _page = null;
    }

    public async ValueTask DisposeAsync() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        HtmlProcessRuntimeSession? page;
        lock (_sync) { page = _page; _page = null; }
        if (page != null) await page.DisposeAsync().ConfigureAwait(false);
    }
}
