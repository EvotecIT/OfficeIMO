using OfficeIMO.Internal;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private sealed partial class WorkflowInputSnapshots {
        private readonly List<(OfficeStreamFileSnapshot Snapshot, OfficeWorkflowStreamInput Source)> _packageSnapshots = [];

        private async Task<ValidatedRequest> CapturePackageAsync(ValidatedRequest request, CancellationToken token) {
            OfficeWorkflowDirectoryPackageInput package = request.InputDirectoryPackage!;
            var sourceGuard = new ProviderPackagePublicationGuard(package.SourcePublicationGuard, null, request.InputPath);
            if (!await sourceGuard.CanPublishAsync(request.OutputPath!, false, token).ConfigureAwait(false))
                throw new IOException("The output is not separate from the source directory package.");
            var captured = await CaptureDirectoryAsync(package.Directory, true, package.MaximumEntries,
                request.Limits.MaximumInputBytes, token).ConfigureAwait(false);
            if (!await sourceGuard.CanPublishAsync(request.OutputPath!, false, token).ConfigureAwait(false))
                throw new IOException("The source directory package changed during capture.");

            // The format owner reads only the private member snapshot. Provider access remains in the shared layer.
            OfficeWorkflowStreamInput source = request.Registration!.DirectoryPackageInput!(captured.Path,
                request.Limits.CloneAndValidate(), request.RegisteredConversionSettings);
            if (source is null || source.SnapshotKind != OfficeWorkflowSourceSnapshotKind.DirectoryPackage || source.SourcePublicationGuard is null)
                throw new ArgumentException("A directory-package owner must provide a directory snapshot input and package output-separation guard.");
            source = new(package.Name, source.OpenRead, source.ExpectedSha256, source.SnapshotKind, source.SourcePublicationGuard);
            var transport = await OfficeStreamFileSnapshot.CaptureAsync(source.OpenRead, Path.GetExtension(package.Name),
                request.Limits.MaximumInputBytes, source.ExpectedSha256, token).ConfigureAwait(false);
            _packageSnapshots.Add((transport, source));
            IOfficeWorkflowPublicationGuard host = new ProviderPackagePublicationGuard(package.SourcePublicationGuard,
                request.PublicationGuard, request.InputPath);
            return request with {
                InputPath = transport.FilePath,
                InputStream = source,
                PublicationGuard = Guard(host, request.Limits.MaximumInputBytes,
                    _directoryFiles.Select(item => item.Access.Location).ToArray(), request.OutputStream)
            };
        }
    }

    /// <summary>Rechecks provider-owned root identity and separation on both sides of host authorization.</summary>
    private sealed class ProviderPackagePublicationGuard(IOfficeWorkflowPublicationGuard source,
        IOfficeWorkflowPublicationGuard? host, string location) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (string.Equals(OfficeStorageIdentity.Normalize(location), OfficeStorageIdentity.Normalize(path), StringComparison.Ordinal)) return false;
            if (!await source.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            token.ThrowIfCancellationRequested();
            return await source.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false);
        }
    }
}
