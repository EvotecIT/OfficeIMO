using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeCloudFontFiles {
    private static readonly object CatalogLock = new();
    private static readonly Dictionary<string, FamilyCatalog> Catalogs = new(StringComparer.Ordinal);

    // Folder names are a fast hint. Localized family aliases need the name tables,
    // indexed once per cache root rather than full font reads for every face request.
    private static IEnumerable<string> GetAliasFiles(string cloudRoot, string familyKey, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!IsOrdinaryDirectory(cloudRoot)) yield break;
        FamilyCatalog? catalog;
        DateTime now = DateTime.UtcNow;
        lock (CatalogLock) Catalogs.TryGetValue(cloudRoot, out catalog);
        if (catalog == null || now - catalog.CreatedUtc > TimeSpan.FromMinutes(1)) {
            var files = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            int inspected = 0, aliases = 0;
            foreach (string family in EnumerateChildren(cloudRoot, directories: true, maximum: 512)) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (string file in EnumerateFontFiles(family)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++inspected > 2048) break;
                    foreach (string key in OfficeFontNameMetadata.ReadFamilyKeys(file, cancellationToken)) {
                        if (++aliases > 8192) break;
                        if (!files.TryGetValue(key, out List<string>? paths)) files[key] = paths = new List<string>();
                        paths.Add(file);
                    }
                    if (aliases > 8192) break;
                }
                if (inspected > 2048 || aliases > 8192) break;
            }
            catalog = new FamilyCatalog(now, files);
            lock (CatalogLock) {
                cancellationToken.ThrowIfCancellationRequested();
                if (Catalogs.Count >= 8) Catalogs.Clear();
                Catalogs[cloudRoot] = catalog;
            }
        }
        if (catalog.Files.TryGetValue(familyKey, out List<string>? found)) {
            foreach (string path in found) {
                cancellationToken.ThrowIfCancellationRequested();
                if (IsOrdinaryDirectory(System.IO.Path.GetDirectoryName(path)!) && IsOrdinaryFile(path)) yield return path;
            }
        }
    }

    private static bool IsOrdinaryFile(string path) {
        try { return (System.IO.File.GetAttributes(path) & (System.IO.FileAttributes.ReparsePoint | System.IO.FileAttributes.Directory)) == 0; }
        catch (Exception exception) when (IsPathFailure(exception)) { return false; }
    }

    private sealed class FamilyCatalog {
        internal FamilyCatalog(DateTime createdUtc, Dictionary<string, List<string>> files) {
            CreatedUtc = createdUtc;
            Files = files;
        }
        internal DateTime CreatedUtc { get; }
        internal Dictionary<string, List<string>> Files { get; }
    }
}
