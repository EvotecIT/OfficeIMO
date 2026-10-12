using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private async Task<int> ExportImagesAsync(PdfDocument snapshot, string destination, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress) {
        // A provider folder (Files on iPad, a sandboxed Mac folder) grants no file-system access of its own: images are
        // staged in the app's temporary storage and each one is published as a new child of the selected folder.
        bool provider = _storage.UsesProviderPublication(destination);
        string? localFolder = null;
        if (!provider) {
            localFolder = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(destination)
                ?? throw new IOException("Choose a folder on this computer to save the images.");
            Directory.CreateDirectory(localFolder);
        }
        string folderName = localFolder ?? _storage.Describe(destination).Name;
        string stagingDirectory = OfficeIMO.Core.Internal.OfficeTemporaryDirectory.Create(
            ".officeimo-image-export-", localFolder);
        async Task<bool> IsTakenAsync(string name) => provider
            ? await _storage.FolderContainsAsync(destination, name, cancellationToken).ConfigureAwait(false)
            : File.Exists(System.IO.Path.Combine(localFolder!, name)) || Directory.Exists(System.IO.Path.Combine(localFolder!, name));
        var staged = new List<(string Path, int PageNumber, string Extension)>();
        try {
            // The snapshot owns one immutable PDF buffer. Visit writes each extracted image to
            // disk before the next one is decoded; cancellation remains attached to the worker.
            await RunNonDetachableCpuWorkAsync(() => {
                snapshot.Images.Visit(image => {
                    if (!image.IsImageFile) return;
                    cancellationToken.ThrowIfCancellationRequested();
                    int index = staged.Count + 1;
                    string path = System.IO.Path.Combine(stagingDirectory, index.ToString(CultureInfo.InvariantCulture) + ".part");
                    using (var stream = OfficeIMO.Core.Internal.OfficeTemporaryFile.CreateAtPath(
                        path, 64 * 1024, FileOptions.SequentialScan))
                        image.CopyTo(stream, cancellationToken);
                    string extension = string.IsNullOrWhiteSpace(image.FileExtension)
                        ? ".bin" : "." + image.FileExtension.TrimStart('.');
                    staged.Add((path, image.PageNumber, extension));
                }, cancellationToken);
                return staged.Count;
            }, cancellationToken).ConfigureAwait(false);
            if (staged.Count == 0) throw new InvalidOperationException("This document has no embedded images that can be saved as image files.");

            string baseName = OfficeIMO.Core.Internal.OfficePortableFileName.SanitizeBaseName(
                System.IO.Path.GetFileNameWithoutExtension(FileName), maximumLength: 80);
            if (baseName.Length == 0) baseName = "document";
            string[] names = new string[staged.Count];
            for (int batch = 0; ; batch++) {
                if (batch == int.MaxValue) throw new IOException("No available image export names remain in this folder.");
                string prefix = batch == 0 ? baseName : string.Create(CultureInfo.InvariantCulture, $"{baseName}-{batch}");
                for (int index = 0; index < staged.Count; index++) {
                    names[index] = string.Create(CultureInfo.InvariantCulture,
                        $"{prefix}-p{staged[index].PageNumber}-{index + 1}{staged[index].Extension}");
                }
                bool free = true;
                foreach (string name in names) {
                    if (!await IsTakenAsync(name).ConfigureAwait(false)) continue;
                    free = false;
                    break;
                }
                if (free) break;
            }
            int published = 0;
            try {
                for (int index = 0; index < staged.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string stagedPath = staged[index].Path;
                    string target = provider
                        ? await _storage.SelectNewFileAsync(destination, names[index], cancellationToken).ConfigureAwait(false)
                        : System.IO.Path.Combine(localFolder!, names[index]);
                    await WriteWorkspaceOutputAsync(target,
                        async (stream, token) => {
                            using var source = new FileStream(stagedPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            await source.CopyToAsync(stream, 64 * 1024, token).ConfigureAwait(false);
                        }, cancellationToken,
                        conflictPolicy: OfficeIMO.Core.Internal.OfficeFileCommit.ConflictPolicy.FailIfExists).ConfigureAwait(false);
                    published++;
                    File.Delete(stagedPath);
                    progress?.Report(new PdfWorkspaceProgress($"Saved {published} of {staged.Count} images", published / (double)staged.Count));
                }
            } catch (Exception error) when (published > 0 &&
                                            error is IOException or UnauthorizedAccessException or OperationCanceledException) {
                throw new IOException($"Image export stopped after saving {published} of {staged.Count} images in '{folderName}'. " +
                    "The saved images remain in that folder; check them before retrying.", error);
            }
            return staged.Count;
        } finally {
            Directory.Delete(stagingDirectory, recursive: true);
        }
    }
}
