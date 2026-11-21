using OfficeIMO.Core.Internal;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private const string ModifiedDateNote = "EPUB 3 requires a modification date; saving an edited publication replaces it with the save time.";

    /// <summary>Finds identity metadata and image metadata and provenance without changing the publication.</summary>
    public OfficeHiddenDataReport InspectHiddenData() => CollectHiddenData(null);

    /// <summary>
    /// Removes the selected categories of metadata and hidden data, then reports what was found and what happened to each item.
    /// Title, language and identifier stay because a valid publication requires them. Save the publication to keep the result.
    /// </summary>
    /// <param name="options">Categories to remove; defaults to <see cref="OfficeHiddenDataKind.Default"/>, which keeps provenance.</param>
    public OfficeHiddenDataReport RemoveHiddenData(OfficeHiddenDataOptions? options = null) =>
        CollectHiddenData(options ?? new OfficeHiddenDataOptions());

    private OfficeHiddenDataReport CollectHiddenData(OfficeHiddenDataOptions? removal) {
        var collector = new OfficeHiddenDataCollector("EPUB", removal);
        CollectEpubDocumentProperties(collector);
        CollectEpubCustomProperties(collector);
        CollectEpubImagesAndProvenance(collector);
        return collector.Build();
    }

    private void CollectEpubDocumentProperties(OfficeHiddenDataCollector collector) {
        const OfficeHiddenDataKind kind = OfficeHiddenDataKind.DocumentProperties;
        string location = PackagePath;
        foreach (XElement element in RequireSection("metadata").Elements().ToList()) {
            if (element.Name.Namespace == Dc) {
                string? description = element.Name.LocalName switch {
                    "creator" => "Creator",
                    "contributor" => "Contributor",
                    "publisher" => "Publisher",
                    "date" => "Date",
                    _ => null
                };
                if (description == null) continue;
                string? role = (string?)element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName is "role" or "event");
                collector.Add(kind, location, role == null ? description : description + " (" + role + ")", element.Value,
                    () => RemoveMetadataElement(element));
            } else if (element.Name == Opf + "meta" && element.Attribute("refines") == null &&
                       EpubVocabulary.Expand(Root, (string?)element.Attribute("property") ?? string.Empty) == "http://purl.org/dc/terms/modified") {
                collector.Add(kind, location, "Modification date", element.Value, null, ModifiedDateNote);
            }
        }
    }

    /// <summary>Named EPUB 2 meta values, such as those producing tools add. The cover declaration is structure and stays.</summary>
    private void CollectEpubCustomProperties(OfficeHiddenDataCollector collector) {
        foreach (XElement element in RequireSection("metadata").Elements(Opf + "meta").ToList()) {
            string? name = (string?)element.Attribute("name");
            if (string.IsNullOrEmpty(name) || name == "cover") continue;
            collector.Add(OfficeHiddenDataKind.CustomProperties, PackagePath, "Metadata '" + name + "'", (string?)element.Attribute("content"),
                () => RemoveMetadataElement(element));
        }
    }

    /// <summary>Removes a metadata element with the refinements that describe it.</summary>
    private void RemoveMetadataElement(XElement element) {
        XElement metadata = RequireSection("metadata");
        int index = metadata.Elements().ToList().IndexOf(element);
        if (index < 0) return;
        string? id = (string?)element.Attribute("id");
        EditPackageElement(metadata, section => {
            XElement target = section.Elements().ElementAt(index);
            if (id != null) section.Elements(Opf + "meta").Where(meta => (string?)meta.Attribute("refines") == "#" + id).Remove();
            target.Remove();
        });
    }

    /// <summary>Provenance and image metadata in the package and its images, through the shared asset engine.</summary>
    private void CollectEpubImagesAndProvenance(OfficeHiddenDataCollector collector) {
        const string manifestPath = OfficeHiddenDataAssets.PackageManifestPath;
        if (_entries.TryGetValue(manifestPath, out byte[]? credentials) &&
            OfficeHiddenDataAssets.ProcessPackageManifest(credentials, manifestPath, collector)) {
            foreach (EpubManifestItem item in Manifest.Where(item => item.Reference.ContainerPath == manifestPath).ToList()) {
                EditPackageElement(RequireSection("manifest"), section => section.Elements(Opf + "item")
                    .Where(element => (string?)element.Attribute("id") == item.Id).Remove());
            }
            _entries.Remove(manifestPath);
            _retainedBytes -= credentials.LongLength;
            _entryOrigins.Remove(manifestPath);
            _mergedEntryOrigins.Remove(manifestPath);
            MarkChanged();
        }

        var processed = new HashSet<string>(StringComparer.Ordinal);
        foreach (EpubManifestItem item in Manifest) {
            if (!item.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) || item.Reference.Kind != EpubReferenceKind.Container) continue;
            string? path = item.Reference.ContainerPath;
            if (path == null || !processed.Add(path) || !_entries.TryGetValue(path, out byte[]? data)) continue;
            if (_encryption.Any(encryption => encryption.Path == path)) continue;
            byte[]? replacement = OfficeHiddenDataAssets.ProcessImage(data, path, collector);
            if (replacement != null) ReplaceResourcePayload(path, replacement);
        }
    }
}
