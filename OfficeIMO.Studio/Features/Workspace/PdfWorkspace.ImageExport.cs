using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private async Task<int> ExportImagesAsync(PdfDocument snapshot, string destination, CancellationToken cancellationToken) {
        destination = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(destination)
            ?? throw new IOException("Choose a folder on this computer to save the images.");
        Directory.CreateDirectory(destination);
        string stagingDirectory = System.IO.Path.Combine(destination, ".officeimo-image-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stagingDirectory);
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
                    using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        image.CopyTo(stream, cancellationToken);
                    staged.Add((path, image.PageNumber, image.FileExtension ?? ".bin"));
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
                if (names.All(name => !File.Exists(System.IO.Path.Combine(destination, name)) &&
                                      !Directory.Exists(System.IO.Path.Combine(destination, name)))) break;
            }
            for (int index = 0; index < staged.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                string stagedPath = staged[index].Path;
                await WriteWorkspaceOutputAsync(System.IO.Path.Combine(destination, names[index]),
                    async (stream, token) => {
                        using var source = new FileStream(stagedPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await source.CopyToAsync(stream, 64 * 1024, token).ConfigureAwait(false);
                    }, cancellationToken,
                    conflictPolicy: OfficeIMO.Core.Internal.OfficeFileCommit.ConflictPolicy.FailIfExists).ConfigureAwait(false);
            }
            return staged.Count;
        } finally {
            Directory.Delete(stagingDirectory, recursive: true);
        }
    }
}
