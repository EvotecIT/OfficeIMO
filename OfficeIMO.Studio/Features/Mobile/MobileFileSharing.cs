using OfficeIMO.Core.Internal;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Keeps a provider-backed output readable for the lifetime of the native share sheet.</summary>
internal static class MobileFileSharing {
    internal static async Task ShareAsync(StudioApplicationServices services, string location, Func<string, Task> share) {
        if (!services.Storage.UsesProviderPublication(location)) {
            await share(location);
            return;
        }
        StudioStorageSnapshot snapshot = await services.Storage.ReadSnapshotAsync(location, CancellationToken.None);
        string root = services.LocalDocuments?.Path ?? throw new IOException("The mobile document folder is unavailable.");
        string directory = Path.Combine(root, ".sharing", Guid.NewGuid().ToString("N"));
        string copy = Path.Combine(directory, Path.GetFileName(services.Storage.Describe(location).Name));
        try {
            Directory.CreateDirectory(directory);
            await Task.Run(() => OfficeFileCommit.WriteAllBytes(copy, snapshot.Bytes, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly));
            await share(copy);
        } finally {
            if (File.Exists(copy)) File.Delete(copy);
            if (Directory.Exists(directory)) Directory.Delete(directory);
        }
    }
}
