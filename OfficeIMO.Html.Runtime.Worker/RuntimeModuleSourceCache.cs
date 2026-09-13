using System.Text;

namespace OfficeIMO.Html.Runtime.Worker;

// Fetching, identity, integrity requirements and source-count limits are kept
// outside Jint so preload and future module engines can share the same owner.
internal sealed class RuntimeModuleSourceCache(RuntimeResourceLoader resources, int maximum, int maximumIntegrityCharacters,
    RuntimeSubresourceIntegrity integrity) {
    private readonly Dictionary<string, Entry> _entries = new(StringComparer.Ordinal);
    private readonly object _sync = new();

    internal Task<RuntimeModuleSource> Register(string identity, RuntimeModuleSource source, string? integrityMetadata) {
        lock (_sync) {
            if (!_entries.TryGetValue(identity, out var entry)) {
                Reserve();
                _entries.Add(identity, entry = new(Task.FromResult(source), integrity, maximumIntegrityCharacters));
            }
            return entry.ReadAsync(integrityMetadata, CancellationToken.None);
        }
    }

    internal Task<RuntimeModuleSource> GetOrLoad(string identity, Uri url, string? integrityMetadata, CancellationToken token = default) {
        Entry entry;
        Task<RuntimeModuleSource> pending;
        lock (_sync) {
            if (!_entries.TryGetValue(identity, out entry!)) {
                Reserve();
                var cancellation = new CancellationTokenSource();
                entry = new(LoadAsync(url, cancellation.Token), integrity, maximumIntegrityCharacters, cancellation);
                _entries.Add(identity, entry);
            }
            pending = entry.ReadAsync(integrityMetadata, token);
        }
        return ReadAndReleaseCanceledAsync(identity, entry, pending, token);
    }

    private async Task<RuntimeModuleSource> ReadAndReleaseCanceledAsync(string identity, Entry entry,
        Task<RuntimeModuleSource> pending, CancellationToken token) {
        try { return await pending.ConfigureAwait(false); }
        catch (OperationCanceledException) when (token.IsCancellationRequested) {
            bool cancel = false;
            lock (_sync) {
                if (_entries.TryGetValue(identity, out var current) && ReferenceEquals(current, entry) && entry.CanCancel) {
                    _entries.Remove(identity);
                    cancel = true;
                }
            }
            if (cancel) entry.Cancel();
            throw;
        }
    }

    private async Task<RuntimeModuleSource> LoadAsync(Uri url, CancellationToken token) {
        var resource = await resources.FetchAsync(url, new RuntimeFetchRequest(), token).ConfigureAwait(false);
        RuntimeModuleLoader.Validate(resource.StatusCode, resource.ContentType);
        string source = Encoding.UTF8.GetString(resource.Buffer);
        if (source.StartsWith('\uFEFF')) source = source[1..];
        return new(source, resource.FinalUrl.AbsoluteUri, resource.Buffer, resource.ContentType,
            resource.StatusCode, resource.Headers);
    }

    private void Reserve() {
        if (_entries.Count >= maximum) throw new HtmlScriptRuntimeException("The module source count budget was exceeded.");
    }

    private sealed class Entry(Task<RuntimeModuleSource> source, RuntimeSubresourceIntegrity integrity,
        int maximumIntegrityCharacters, CancellationTokenSource? cancellation = null) {
        private readonly HashSet<string> _requirements = new(StringComparer.Ordinal);
        private Exception? _failure;
        private int _readers;
        private int _requirementCharacters;

        internal bool CanCancel => Volatile.Read(ref _readers) == 0 && !source.IsCompleted;

        internal void Cancel() => cancellation?.Cancel();

        internal void RequireIntegrity(string? metadata) {
            if (string.IsNullOrWhiteSpace(metadata)) return;
            lock (_requirements) {
                if (_failure != null) throw _failure;
                if (_requirements.Contains(metadata)) return;
                if (metadata.Length > maximumIntegrityCharacters - _requirementCharacters)
                    throw _failure = new HtmlScriptRuntimeException("Module integrity metadata exceeds its per-source character budget.");
                _requirements.Add(metadata);
                _requirementCharacters += metadata.Length;
            }
        }

        internal Task<RuntimeModuleSource> ReadAsync(string? metadata, CancellationToken token) {
            RequireIntegrity(metadata);
            Interlocked.Increment(ref _readers);
            return ReadCoreAsync(token);
        }

        private async Task<RuntimeModuleSource> ReadCoreAsync(CancellationToken token) {
            try {
                RuntimeModuleSource value = await source.WaitAsync(token).ConfigureAwait(false);
                lock (_requirements) {
                    if (_failure != null) throw _failure;
                    foreach (string requirement in _requirements) {
                        if (integrity.IsSatisfied(value.Buffer, requirement)) continue;
                        throw _failure = new HtmlScriptRuntimeException("Module subresource integrity verification failed.");
                    }
                }
                return value;
            } finally { Interlocked.Decrement(ref _readers); }
        }
    }
}

internal sealed record RuntimeModuleSource(string Text, string Location, byte[] Buffer, string ContentType,
    int StatusCode, IReadOnlyDictionary<string, string> Headers);
