namespace OfficeIMO.Web.Converter.Services;

/// <summary>Owns temporary download URLs until a current operation publishes all of its results.</summary>
internal sealed class ConverterObjectUrlBatch(ConverterInterop interop, Func<bool> isCurrent) : IAsyncDisposable {
    private readonly List<string> _urls = [];
    private bool _committed;

    internal async ValueTask<string> CreateAsync(byte[] bytes, string contentType) {
        EnsureCurrent();
        string url = await interop.CreateObjectUrlAsync(bytes, contentType);
        _urls.Add(url);
        EnsureCurrent();
        return url;
    }

    internal void Commit() {
        EnsureCurrent();
        _committed = true;
    }

    private void EnsureCurrent() {
        if (!isCurrent()) throw new OperationCanceledException("The workspace changed while creating the result.");
    }

    public async ValueTask DisposeAsync() {
        if (_committed) return;
        foreach (string url in _urls) await interop.RevokeObjectUrlAsync(url);
    }
}
