using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Keeps local provider access alive while inspecting filesystem identity and host ownership.</summary>
    private sealed class OcrScopedSourcePublicationGuard(IOfficeWorkflowPublicationGuard? host,
        string source, OfficeWorkflowStreamInput input, OfficeWorkflowStreamOutput? output) : IOfficeWorkflowPublicationGuard {
        private readonly string? _localSource = OfficeStorageIdentity.GetLocalPath(source);
        private string? _identity;

        internal async Task<Stream> OpenReadAsync(CancellationToken token) {
            Stream stream = await input.OpenRead(token).ConfigureAwait(false);
            try {
                token.ThrowIfCancellationRequested();
                if (_localSource is not null) {
                    string identity = stream is FileStream file
                        ? OfficePathIdentity.GetPhysicalIdentityKey(_localSource, file.SafeFileHandle)
                        : OfficePathIdentity.GetPhysicalIdentityKey(_localSource);
                    if (_identity is not null && identity != _identity) throw new IOException("The source PDF was replaced during OCR.");
                    if (OfficePathIdentity.GetPhysicalIdentityKey(_localSource) != identity) throw new IOException("The source PDF changed while opening provider access.");
                    _identity ??= identity;
                }
                return stream;
            } catch {
                await stream.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            // Opening a provider stream can grant the only permission to inspect its local path.
            await using Stream? sourceAccess = _localSource is null ? null : await OpenReadAsync(token).ConfigureAwait(false);
            await using Stream? outputAccess = await OpenOutputAccessAsync(path, token).ConfigureAwait(false);
            if (!SourceIsSeparate(path)) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            token.ThrowIfCancellationRequested();
            return SourceIsSeparate(path);
        }

        private bool SourceIsSeparate(string path) {
            if (_localSource is not null && OfficePathIdentity.GetPhysicalIdentityKey(_localSource) != _identity)
                throw new IOException("The source PDF was replaced during OCR.");
            return !OfficeStorageIdentity.AreEquivalent(source, path);
        }

        private async Task<Stream?> OpenOutputAccessAsync(string path, CancellationToken token) {
            if (output is null || OfficeStorageIdentity.GetLocalPath(path) is null) return null;
            try { return await output.OpenRead(token).ConfigureAwait(false); }
            catch (FileNotFoundException) { return null; } // A selected new file has no existing identity to protect.
        }
    }
}
