namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private void EnsureResourceMutationAllowed(string path, bool removing = false) {
        // Signature removal is finalized only by the explicit policy at Write/Save.
        if (removing && path == "META-INF/signatures.xml") return;
        if (path == "mimetype" || _rootfilePaths.Contains(path, StringComparer.Ordinal) ||
            path.StartsWith("META-INF/", StringComparison.Ordinal))
            throw new InvalidOperationException("Container controls and declared rootfiles cannot be changed through raw resource APIs.");
    }

    // A selected-package removal must also leave retained renditions usable.
    private void EnsureRemovalPreservesRootfiles(string removedPath) {
        if (_rootfilePaths.Contains(removedPath, StringComparer.Ordinal))
            throw new InvalidOperationException("A declared rootfile cannot be removed as a resource.");
        foreach (string rootfilePath in _rootfilePaths.Where(path => path != PackagePath)) {
            try {
                if (!_entries.TryGetValue(rootfilePath, out byte[]? payload))
                    throw new InvalidDataException("Retained rootfile is missing.");
                XDocument package = ParseXml(payload, Math.Min(_maximumMetadataBytes, _maximumEntryBytes));
                XElement root = package.Root ?? throw new InvalidDataException("Retained rootfile has no package root.");
                XElement manifest = root.Element(Opf + "manifest") ?? throw new InvalidDataException("Retained rootfile has no manifest.");
                if (root.Name != Opf + "package") throw new InvalidDataException("Retained rootfile is not an OPF package.");
                EpubManifestItem[] items = manifest.Elements(Opf + "item").Select(item => new EpubManifestItem(item, rootfilePath)).ToArray();
                if (items.Any(item => item.Reference.ContainerPath == removedPath) ||
                    PackageResourceReferences(root).Any(attribute => EpubReference.Resolve(rootfilePath, attribute.Value).ContainerPath == removedPath))
                    throw new InvalidOperationException("Resource is referenced by retained rootfile " + rootfilePath + ".");
                foreach (EpubManifestItem item in items.Where(item => item.Reference.Kind == EpubReferenceKind.Container &&
                    (HasMediaType(item.MediaType, "application/xhtml+xml") || HasMediaType(item.MediaType, "image/svg+xml") ||
                        HasMediaType(item.MediaType, "application/x-dtbncx+xml")))) {
                    string path = RequireLocalPath(item);
                    if (_encryption.Any(encryption => encryption.Path == path && encryption.RequiresDecryption) ||
                        !_entries.TryGetValue(path, out byte[]? contentBytes))
                        throw new InvalidDataException("Retained rendition content cannot be inspected.");
                    XDocument content = ParseXml(contentBytes, _maximumEntryBytes);
                    if (ContentResourceReferences(content, path).Any(reference => reference.ContainerPath == removedPath))
                        throw new InvalidOperationException("Resource is referenced by retained rendition content " + path + ".");
                }
            } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException) {
                throw new InvalidOperationException("Cannot safely remove a resource while retained rootfile " + rootfilePath + " is unreadable.", exception);
            }
        }
    }
}
