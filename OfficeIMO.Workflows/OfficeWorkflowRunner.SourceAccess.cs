using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Captures and verifies local identity inside the selected provider's readable access scope.</summary>
    private sealed class WorkflowSourceAccess(string location, OfficeWorkflowStreamInput source) {
        internal string Location { get; } = location;
        internal string? LocalPath { get; } = OfficeStorageIdentity.GetLocalPath(location);
        private string? _identity;

        internal OfficeWorkflowStreamInput CreateInput() => new(source.Name, OpenReadAsync, source.ExpectedSha256);

        internal async Task<Stream> OpenReadAsync(CancellationToken token) {
            Stream stream = await source.OpenRead(token).ConfigureAwait(false);
            try {
                token.ThrowIfCancellationRequested();
                if (LocalPath is not null) {
                    string identity = stream is FileStream file
                        ? OfficePathIdentity.GetPhysicalIdentityKey(LocalPath, file.SafeFileHandle)
                        : OfficePathIdentity.GetPhysicalIdentityKey(LocalPath);
                    if (_identity is not null && identity != _identity) throw new IOException("The workflow source was replaced during execution.");
                    if (OfficePathIdentity.GetPhysicalIdentityKey(LocalPath) != identity)
                        throw new IOException("The workflow source changed while opening provider access.");
                    _identity ??= identity;
                }
                return stream;
            } catch {
                await stream.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        internal void VerifyIdentity() {
            if (LocalPath is not null && OfficePathIdentity.GetPhysicalIdentityKey(LocalPath) != _identity)
                throw new IOException("The workflow source was replaced during execution.");
        }
    }

    /// <summary>Holds provider scopes through host authorization and both source/output separation checks.</summary>
    private sealed class WorkflowScopedSourcePublicationGuard : IOfficeWorkflowPublicationGuard {
        private readonly IOfficeWorkflowPublicationGuard? _host;
        private readonly string[] _sources;
        private readonly WorkflowSourceAccess[] _accesses;
        private readonly WorkflowSourceAccess[] _localScopes;
        private readonly (string Path, string Identity)[] _localSources;
        private readonly OfficeWorkflowStreamOutput? _output;

        internal WorkflowScopedSourcePublicationGuard(IOfficeWorkflowPublicationGuard? host, string[] sources,
            WorkflowSourceAccess[] accesses, OfficeWorkflowStreamOutput? output) {
            _host = host;
            _sources = sources;
            _accesses = accesses;
            _output = output;
            _localScopes = accesses.Where(access => access.LocalPath is not null)
                .GroupBy(access => access.Location, StringComparer.Ordinal).Select(group => group.First()).ToArray();
            _localSources = sources.Where(source => !accesses.Any(access => access.Location == source))
                .Select(OfficeStorageIdentity.GetLocalPath).OfType<string>()
                .Select(path => (path, OfficePathIdentity.GetPhysicalIdentityKey(path))).ToArray();
        }

        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            var scopes = new List<Stream>(_localScopes.Length + 1);
            try {
                foreach (var access in _localScopes) scopes.Add(await access.OpenReadAsync(token).ConfigureAwait(false));
                if (_output is not null && OfficeStorageIdentity.GetLocalPath(path) is not null) {
                    try { scopes.Add(await _output.OpenRead(token).ConfigureAwait(false)); }
                    catch (FileNotFoundException) { } // A selected new output has no existing file identity.
                }
                token.ThrowIfCancellationRequested();
                if (!SourcesAreSeparate(path, isDirectory)) return false;
                if (_host is not null && !await _host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
                token.ThrowIfCancellationRequested();
                return SourcesAreSeparate(path, isDirectory);
            } finally {
                List<Exception>? failures = null;
                for (int index = scopes.Count - 1; index >= 0; index--) {
                    try { await scopes[index].DisposeAsync().ConfigureAwait(false); }
                    catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                        (failures ??= new()).Add(error);
                    }
                }
                if (failures is not null) throw new IOException("Provider access scopes could not be released.", new AggregateException(failures));
            }
        }

        private bool SourcesAreSeparate(string path, bool isDirectory) {
            foreach (var access in _accesses) access.VerifyIdentity();
            foreach (var source in _localSources) {
                if (OfficePathIdentity.GetPhysicalIdentityKey(source.Path) != source.Identity)
                    throw new IOException("The workflow source was replaced during execution.");
            }
            if (_sources.Any(source => OfficeStorageIdentity.AreEquivalent(source, path))) return false;
            string? outputDirectory = isDirectory ? OfficeStorageIdentity.GetLocalPath(path) : null;
            return outputDirectory is null || !_sources.Any(source => OfficeStorageIdentity.GetLocalPath(source) is { } local &&
                OfficePathIdentity.IsSameOrDescendant(local, outputDirectory));
        }
    }
}
