using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Compares this baseline with a revised edition without mutation. Resources match by manifest
    /// ID; unmanifested payloads match by path. XML comparison is structural, not rendered equivalence.
    /// </summary>
    public EpubEditionComparison CompareTo(EpubPublication revised, CancellationToken cancellationToken = default) {
        if (revised == null) throw new ArgumentNullException(nameof(revised));
        cancellationToken.ThrowIfCancellationRequested();
        var changes = new List<EpubEditionResourceChange>();
        var textChanges = new List<EpubEditionTextChange>();
        // Materialize implicit namespace declarations as XML serialization does, without writing a ZIP.
        var oldPackage = ParseXml(SerializeXml(_package, _maximumMetadataBytes), _maximumMetadataBytes);
        var newPackage = ParseXml(SerializeXml(revised._package, revised._maximumMetadataBytes), revised._maximumMetadataBytes);
        var oldItems = Manifest.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var newItems = revised.Manifest.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var oldDeclarations = oldPackage.Root!.Element(Opf + "manifest")!.Elements(Opf + "item").ToDictionary(item => (string)item.Attribute("id")!, StringComparer.Ordinal);
        var newDeclarations = newPackage.Root!.Element(Opf + "manifest")!.Elements(Opf + "item").ToDictionary(item => (string)item.Attribute("id")!, StringComparer.Ordinal);
        var coveredOld = new HashSet<string>(StringComparer.Ordinal) { PackagePath };
        var coveredNew = new HashSet<string>(StringComparer.Ordinal) { revised.PackagePath };
        foreach (string id in oldItems.Keys.Union(newItems.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            oldItems.TryGetValue(id, out EpubManifestItem? before); newItems.TryGetValue(id, out EpubManifestItem? after);
            string? oldPath = before?.Reference.ContainerPath, newPath = after?.Reference.ContainerPath;
            if (oldPath != null) coveredOld.Add(oldPath); if (newPath != null) coveredNew.Add(newPath);
            EpubEditionChangeKind kind = before == null ? EpubEditionChangeKind.Added : after == null ? EpubEditionChangeKind.Removed : EpubEditionChangeKind.None;
            if (before != null && after != null) {
                if (before.Reference.ResolvedValue != after.Reference.ResolvedValue) kind |= EpubEditionChangeKind.Location;
                XElement oldDeclaration = oldDeclarations[id];
                XElement newDeclaration = newDeclarations[id];
                if (!XNode.DeepEquals(oldDeclaration, newDeclaration)) kind |= EpubEditionChangeKind.Declaration;
                if (oldPath != null && newPath != null) {
                    byte[] left = _entries[oldPath], right = revised._entries[newPath];
                    if (!left.SequenceEqual(right)) kind |= CompareEditionPayload(before, after, left, right, revised, textChanges, cancellationToken);
                }
            }
            if (kind != EpubEditionChangeKind.None) changes.Add(new EpubEditionResourceChange(id, before?.Reference.ResolvedValue, after?.Reference.ResolvedValue, kind));
        }
        foreach (string path in _entries.Keys.Where(path => !coveredOld.Contains(path)).Union(revised._entries.Keys.Where(path => !coveredNew.Contains(path)), StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            bool oldExists = !coveredOld.Contains(path) && _entries.TryGetValue(path, out _);
            bool newExists = !coveredNew.Contains(path) && revised._entries.TryGetValue(path, out _);
            EpubEditionChangeKind kind = !oldExists ? EpubEditionChangeKind.Added : !newExists ? EpubEditionChangeKind.Removed :
                _entries[path].SequenceEqual(revised._entries[path]) ? EpubEditionChangeKind.None : EpubEditionChangeKind.Binary;
            if (kind != EpubEditionChangeKind.None) changes.Add(new EpubEditionResourceChange(null, oldExists ? path : null, newExists ? path : null, kind));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new EpubEditionComparison(!XNode.DeepEquals(EditionMetadata(oldPackage), EditionMetadata(newPackage)),
            !XNode.DeepEquals(oldPackage.Root!.Element(Opf + "spine"), newPackage.Root!.Element(Opf + "spine")),
            PackagePath != revised.PackagePath || !XNode.DeepEquals(EditionPackageStructure(oldPackage), EditionPackageStructure(newPackage)), changes, textChanges);
    }

    private EpubEditionChangeKind CompareEditionPayload(EpubManifestItem before, EpubManifestItem after, byte[] left, byte[] right, EpubPublication revised, List<EpubEditionTextChange> textChanges, CancellationToken token) {
        string type = before.MediaType.ToLowerInvariant();
        bool xml = type is "application/xhtml+xml" or "image/svg+xml" or "application/x-dtbncx+xml" or "application/smil+xml";
        if (!xml || !string.Equals(before.MediaType, after.MediaType, StringComparison.OrdinalIgnoreCase) ||
            _encryption.Any(item => item.Path == before.Reference.ContainerPath) || revised._encryption.Any(item => item.Path == after.Reference.ContainerPath)) return EpubEditionChangeKind.Binary;
        XDocument oldXml = ParseXml(left, _maximumEntryBytes), newXml = ParseXml(right, revised._maximumEntryBytes);
        if (XNode.DeepEquals(oldXml, newXml)) return EpubEditionChangeKind.Serialization;
        if (type == "application/xhtml+xml") CompareEditionText(before.Id, oldXml, newXml, textChanges, token);
        string oldText = type == "application/xhtml+xml" ? oldXml.Root?.Element(Html + "body")?.Value ?? string.Empty : oldXml.Root?.Value ?? string.Empty;
        string newText = type == "application/xhtml+xml" ? newXml.Root?.Element(Html + "body")?.Value ?? string.Empty : newXml.Root?.Value ?? string.Empty;
        return EpubEditionChangeKind.Xml | (oldText == newText ? EpubEditionChangeKind.None : EpubEditionChangeKind.Text);
    }

    private static XElement? EditionMetadata(XDocument package) {
        XElement? source = package.Root!.Element(Opf + "metadata");
        if (source == null) return null;
        var metadata = new XElement(source);
        metadata.Elements(Opf + "meta").Where(element => element.Attribute("refines") == null &&
            EpubVocabulary.Expand(package.Root, (string?)element.Attribute("property") ?? string.Empty) == "http://purl.org/dc/terms/modified").Remove();
        return metadata;
    }

    private static XDocument EditionPackageStructure(XDocument package) {
        var copy = new XDocument(package);
        copy.Root!.Elements().Where(element => element.Name == Opf + "metadata" || element.Name == Opf + "spine").Remove();
        // Keep manifest scaffolding and extensions, while item declarations are compared by identity.
        copy.Root.Element(Opf + "manifest")?.Elements(Opf + "item").Remove();
        return copy;
    }
}
