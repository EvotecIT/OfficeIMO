namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfPackage {
    // Apply save preparation only after the destination accepted the bytes. Keep
    // content XML nodes in place so existing typed wrappers remain usable.
    internal void AcceptSerializedChanges(OdfPackage output) {
        foreach (OdfPackageEntry entry in _entries) {
            if (!entry.IsRemoved && output._entriesByName.TryGetValue(entry.Name, out OdfPackageEntry? saved) && saved.IsRemoved) {
                entry.Remove();
                _entryGraphChanged = true;
            }
        }
        if (Version != output.Version) {
            UpdateXmlVersions(output.Version);
            _entryGraphChanged = true;
        }
        if (_entryGraphChanged) RebuildManifest(output.Version);
        Version = output.Version;
        _pendingOutputEncrypted = output._pendingOutputEncrypted;
        AcceptChanges();
    }

    internal OdfPackage ForValidation() {
        if (!_entryGraphChanged) return this;
        OdfPackage clone = CloneForSerialization();
        XElement root = clone.GetXml("META-INF/manifest.xml").Root!;
        var listed = new HashSet<string>(root.Elements(OdfNamespaces.Manifest + "file-entry")
            .Select(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") ?? string.Empty), StringComparer.Ordinal);
        var removed = new HashSet<string>(clone._entries.Where(entry => entry.IsRemoved)
            .Select(entry => entry.Name), StringComparer.Ordinal);
        root.Elements(OdfNamespaces.Manifest + "file-entry")
            .Where(element => removed.Contains((string?)element.Attribute(OdfNamespaces.Manifest + "full-path") ?? string.Empty))
            .Remove();
        foreach (OdfPackageEntry entry in clone._entries) {
            if (!entry.IsRemoved && entry.IsNew && !entry.Name.StartsWith("META-INF/", StringComparison.Ordinal) &&
                entry.Name != "mimetype" && listed.Add(entry.Name)) {
                root.Add(OdfPackageTemplates.FileEntry(entry.Name, entry.MediaType ?? GuessMediaType(entry.Name), null));
            }
        }
        return clone;
    }
    // Save preparation changes versions, manifests and signatures. Independent outputs
    // must not apply those changes to the document associated with the source file.
    internal OdfPackage CloneForSerialization(OdfCompatibilityProfile profile = OdfCompatibilityProfile.PreserveSource) {
        EnsureGradientAnglesSurviveVersionChange(ResolveOutputVersion(profile));
        bool rewriteXmlVersions = ResolveOutputVersion(profile) != Version;
        var clone = new OdfPackage(Kind, Version, _loadOptions) {
            _entryGraphChanged = _entryGraphChanged,
            _sourceIsEncrypted = _sourceIsEncrypted,
            ContentEditVersion = ContentEditVersion,
            ExternalXmlEditVersion = ExternalXmlEditVersion,
            StyleLookupVersion = StyleLookupVersion
        };
        foreach (OdfPackageEntry entry in _entries) {
            // Save preparation writes only the manifest unless it changes XML
            // versions. Other trees are read-only during serialization; sharing
            // them avoids copying the complete content graph for ordinary saves.
            OdfPackageEntry copy = entry.CloneForSerialization(rewriteXmlVersions || entry.Name == "META-INF/manifest.xml");
            clone._entries.Add(copy);
            clone._entriesByName.Add(copy.Name, copy);
        }
        clone._diagnostics.AddRange(_diagnostics);
        return clone;
    }
}
