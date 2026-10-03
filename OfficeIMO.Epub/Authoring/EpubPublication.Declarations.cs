namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private bool IsSupportedContentDocument(string mediaType) => HasMediaType(mediaType, "application/xhtml+xml") ||
        (PackageVersion == "3.0" && HasMediaType(mediaType, "image/svg+xml"));

    private static bool IsImageMediaType(string mediaType) => mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);

    private void ValidateManifestDeclarations(EpubManifestItem[] manifest, IReadOnlyDictionary<string, EpubManifestItem> byId, XElement metadata) {
        foreach (EpubManifestItem item in manifest.Where(item => item.MediaOverlayId != null)) {
            EpubManifestItem overlay = byId[item.MediaOverlayId!];
            if (PackageVersion != "3.0" || !IsSupportedContentDocument(item.MediaType) || !HasMediaType(overlay.MediaType, "application/smil+xml"))
                throw new InvalidDataException("A media-overlay association requires an EPUB 3 content document and a SMIL target.");
        }
        EpubManifestItem[] covers = manifest.Where(item => HasToken(item.Properties, "cover-image")).ToArray();
        if (covers.Length > 1 || covers.Any(item => !IsImageMediaType(item.MediaType)))
            throw new InvalidDataException("The cover-image property must select at most one image resource.");
        foreach (XElement cover in metadata.Elements(Opf + "meta").Where(element => (string?)element.Attribute("name") == "cover")) {
            string id = (string?)cover.Attribute("content") ?? string.Empty;
            if (!byId.TryGetValue(id, out EpubManifestItem? target) || !IsImageMediaType(target.MediaType))
                throw new InvalidDataException("Legacy cover metadata must select a manifest image resource.");
        }
    }
}
