using OfficeIMO.Epub;
using OfficeIMO.Reader;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    private static IEnumerable<OfficeDocumentAsset> BuildEpubAssets(EpubDocument document, string sourcePath) {
        int assetIndex = 0;
        foreach (EpubResource resource in document.Resources) {
            string? kind = GetEpubAssetKind(resource.MediaType);
            if (kind == null) continue;

            string id = "epub-" + kind + "-" + assetIndex.ToString("D4", CultureInfo.InvariantCulture);
            string resourceLocation = resource.IsRemote && !string.IsNullOrWhiteSpace(resource.RemoteUri)
                ? RemoveEpubUrlFragment(resource.RemoteUri!)
                : BuildVirtualPath(sourcePath, resource.Path);
            string extensionSource = GetEpubExtensionSource(resource);
            string extension = Path.GetExtension(extensionSource);
            yield return new OfficeDocumentAsset {
                Id = id,
                Kind = kind,
                MediaType = resource.MediaType,
                Extension = string.IsNullOrWhiteSpace(extension) ? null : extension,
                FileName = OfficeDocumentAssetNaming.BuildFileName(id, extension),
                LengthBytes = resource.LengthBytes,
                PayloadHash = resource.Data == null ? null : ComputeEpubPayloadHash(resource.Data),
                PayloadBytes = resource.Data,
                SourceObjectId = resource.Id,
                Location = new ReaderLocation {
                    Path = resourceLocation,
                    SourceBlockKind = kind,
                    BlockAnchor = id
                }
            };
            assetIndex++;
        }
    }

    private static void AddEpubChapterAssets(
        string sourcePath,
        EpubChapter chapter,
        IReadOnlyList<OfficeDocumentAsset> htmlAssets,
        List<OfficeDocumentAsset> documentAssets,
        Dictionary<string, OfficeDocumentAsset> assetsByLocation,
        List<OfficeDocumentAsset> chapterAssets,
        List<OfficeDocumentDiagnostic> diagnostics) {
        foreach (OfficeDocumentAsset htmlAsset in htmlAssets) {
            OfficeDocumentAsset? mappedAsset = null;
            if (htmlAsset.PayloadBytes == null) {
                EpubReference reference = EpubReference.Resolve(chapter.Path, chapter.BaseHref, htmlAsset.SourceObjectId ?? string.Empty);
                if (!reference.IsValid) {
                    AddEpubReferenceDiagnostic(sourcePath, chapter, reference, diagnostics);
                    continue;
                }
                if (reference.Kind == EpubReferenceKind.Data) continue;

                AddEpubReferenceDiagnostic(sourcePath, chapter, reference, diagnostics);
                string? resourceLocation = BuildEpubResolvedLocation(sourcePath, reference, includeFragment: true);
                if (!string.IsNullOrWhiteSpace(resourceLocation)) {
                    mappedAsset = FindEpubAsset(sourcePath, reference, assetsByLocation);
                    if (mappedAsset == null) {
                        htmlAsset.Location.Path = resourceLocation;
                    } else {
                        ApplyEpubOccurrenceMetadata(mappedAsset, htmlAsset);
                    }
                }
            }

            if (mappedAsset == null) {
                mappedAsset = htmlAsset;
                documentAssets.Add(mappedAsset);
                if (!string.IsNullOrWhiteSpace(mappedAsset.Location.Path)
                    && !assetsByLocation.ContainsKey(mappedAsset.Location.Path!)) {
                    assetsByLocation.Add(mappedAsset.Location.Path!, mappedAsset);
                }
            }
            if (!chapterAssets.Contains(mappedAsset)) chapterAssets.Add(mappedAsset);
        }
    }

    private static Dictionary<string, OfficeDocumentAsset> BuildEpubAssetIndex(
        IEnumerable<OfficeDocumentAsset> assets) {
        var result = new Dictionary<string, OfficeDocumentAsset>(StringComparer.Ordinal);
        foreach (OfficeDocumentAsset asset in assets) {
            if (!string.IsNullOrWhiteSpace(asset.Location.Path)
                && !result.ContainsKey(asset.Location.Path!)) {
                result.Add(asset.Location.Path!, asset);
            }
        }
        return result;
    }

    private static void ApplyEpubOccurrenceMetadata(OfficeDocumentAsset packageAsset, OfficeDocumentAsset occurrenceAsset) {
        if (string.IsNullOrWhiteSpace(packageAsset.AltText)
            && !string.IsNullOrWhiteSpace(occurrenceAsset.AltText)) {
            packageAsset.AltText = occurrenceAsset.AltText;
        }
        if (packageAsset.Title == null && occurrenceAsset.Title != null) packageAsset.Title = occurrenceAsset.Title;
        if (!packageAsset.Width.HasValue && occurrenceAsset.Width.HasValue) packageAsset.Width = occurrenceAsset.Width;
        if (!packageAsset.Height.HasValue && occurrenceAsset.Height.HasValue) packageAsset.Height = occurrenceAsset.Height;
    }

    private static string? GetEpubAssetKind(string? mediaType) {
        if (string.IsNullOrWhiteSpace(mediaType)) return "resource";
        string value = mediaType!.Trim();
        if (value.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
            || value.Equals("text/html", StringComparison.OrdinalIgnoreCase)
            || value.Equals("application/x-dtbncx+xml", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }
        if (value.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return "image";
        if (value.StartsWith("audio/", StringComparison.OrdinalIgnoreCase)) return "audio";
        if (value.StartsWith("video/", StringComparison.OrdinalIgnoreCase)) return "video";
        if (value.StartsWith("font/", StringComparison.OrdinalIgnoreCase)
            || value.IndexOf("font", StringComparison.OrdinalIgnoreCase) >= 0
            || value.Equals("application/vnd.ms-opentype", StringComparison.OrdinalIgnoreCase)) {
            return "font";
        }
        if (value.Equals("text/css", StringComparison.OrdinalIgnoreCase)) return "stylesheet";
        if (value.IndexOf("javascript", StringComparison.OrdinalIgnoreCase) >= 0
            || value.Equals("application/ecmascript", StringComparison.OrdinalIgnoreCase)) {
            return "script";
        }
        if (value.Equals("application/smil+xml", StringComparison.OrdinalIgnoreCase)) return "media-overlay";
        return "resource";
    }

    private static string GetEpubExtensionSource(EpubResource resource) {
        string value = resource.IsRemote && !string.IsNullOrWhiteSpace(resource.RemoteUri)
            ? resource.RemoteUri!
            : resource.Path;
        int fragment = value.IndexOf('#');
        if (fragment >= 0) value = value.Substring(0, fragment);
        int query = value.IndexOf('?');
        if (query >= 0) value = value.Substring(0, query);
        if (value.StartsWith("//", StringComparison.Ordinal)) value = "https:" + value;
        return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) ? uri.AbsolutePath : value;
    }

    private static string RemoveEpubUrlFragment(string value) {
        int fragment = value.IndexOf('#');
        return fragment < 0 ? value : value.Substring(0, fragment);
    }

    private static string ComputeEpubPayloadHash(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        return ComputeSha256Hex(stream);
    }
}
