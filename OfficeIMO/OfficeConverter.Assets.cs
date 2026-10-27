using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task CheckAssetsAsync(OfficeConversionAssets assets, OfficeConversionOptions settings,
        string? source, string? output, CancellationToken token) {
        if (assets.Pending.Count == 0) return;
        foreach (OfficeConversionAssets.Asset asset in assets.Pending) {
            if ((source != null && OfficePathIdentity.AreEquivalent(source, asset.TargetPath))
                || (output != null && OfficePathIdentity.AreEquivalent(output, asset.TargetPath)))
                throw new IOException("A conversion sidecar cannot replace the source or main destination.");
            await CheckStagingAsync(settings.PublicationGuard, Path.GetDirectoryName(asset.TargetPath)!, token).ConfigureAwait(false);
            if (settings.PublicationGuard != null && !await settings.PublicationGuard.CanPublishAsync(asset.TargetPath, false, token).ConfigureAwait(false))
                throw new UnauthorizedAccessException("The conversion sidecar was rejected by the publication guard.");
            if (Directory.Exists(asset.TargetPath)) throw new IOException("A conversion sidecar destination is a directory.");
            if (File.Exists(asset.TargetPath) && settings.FileConflictPolicy != OfficePublicationPolicy.Replace)
                throw new IOException("A conversion sidecar already exists. Choose Replace or a new asset directory.");
        }
        if (output != null && File.Exists(output) && settings.FileConflictPolicy == OfficePublicationPolicy.Fail)
            throw new IOException("The conversion destination already exists.");
        if (output != null && settings.PublicationGuard != null
            && !await settings.PublicationGuard.CanPublishAsync(output, false, token).ConfigureAwait(false))
            throw new UnauthorizedAccessException("The conversion destination was rejected by the publication guard.");
    }

    private static async Task PublishAssetAsync(OfficeConversionAssets.Asset asset, OfficeConversionOptions settings, CancellationToken token) {
        string directory = Path.GetDirectoryName(asset.TargetPath)!;
        await CheckStagingAsync(settings.PublicationGuard, directory, token).ConfigureAwait(false);
        Directory.CreateDirectory(directory);
        await CheckStagingAsync(settings.PublicationGuard, directory, token).ConfigureAwait(false);
        string stage = Path.Combine(directory, "." + Path.GetFileName(asset.TargetPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            using (var input = File.OpenRead(asset.StagedPath))
            using (var output = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                await input.CopyToAsync(output, 81920, token).ConfigureAwait(false);
            await OfficeFilePublisher.PublishAsync(stage, asset.TargetPath, settings.FileConflictPolicy,
                settings.PublicationGuard, token).ConfigureAwait(false);
        } finally { if (File.Exists(stage)) File.Delete(stage); }
    }
}
