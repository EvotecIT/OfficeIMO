using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Merges consecutive compatible reflowable chapters into the first resource. Retains both
    /// navigation entries, targets the second chapter's start with boundaryId, and repairs references
    /// atomically. Conflicting scaffolding, styles, identifiers and package refinements are rejected.
    /// </summary>
    public void MergeChapters(string firstManifestId, string secondManifestId, string boundaryId, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        RequireText(boundaryId, nameof(boundaryId)); XmlConvert.VerifyNCName(boundaryId);
        EpubManifestItem firstItem = RequireManifestItem(firstManifestId), secondItem = RequireManifestItem(secondManifestId);
        XElement firstPosition = RequireRestructurablePosition(firstItem), secondPosition = RequireRestructurablePosition(secondItem);
        if (firstPosition.ElementsAfterSelf(Opf + "itemref").FirstOrDefault() != secondPosition)
            throw new InvalidOperationException("Merge chapters in consecutive reading order, first followed by second.");
        VerifyMergeDeclarations(firstItem, secondItem, firstPosition, secondPosition);
        string firstPath = RequireLocalPath(firstItem), secondPath = RequireLocalPath(secondItem);
        EnsureResourceMutationAllowed(firstPath); EnsureResourceMutationAllowed(secondPath, removing: true);
        XDocument first = EditableXhtml(firstManifestId), second = EditableXhtml(secondManifestId);
        VerifyMergeDocument(first); VerifyMergeDocument(second);
        XElement firstBody = first.Root!.Element(Html + "body")!, secondBody = second.Root!.Element(Html + "body")!;
        var firstIds = EpubContentIdentifiers.Collect(first.Root, firstPath, true, cancellationToken);
        var secondIds = EpubContentIdentifiers.Collect(second.Root, secondPath, true, cancellationToken);
        if (firstIds.Contains(boundaryId) || secondIds.Contains(boundaryId)) throw new ArgumentException("The merge boundary ID already exists in chapter content.", nameof(boundaryId));
        if (!SameMergeAttributes(first.Root, second.Root) || !SameMergeAttributes(firstBody, secondBody))
            throw new NotSupportedException("Resolve differing root/body attributes before merging; their language, direction, styling and semantics cannot be discarded.");
        if (!MergeRootNotes(first.Root).SequenceEqual(MergeRootNotes(second.Root), StringComparer.Ordinal))
            throw new NotSupportedException("Chapter root-level annotations conflict.");
        var shared = FindMergeSeam(firstBody, secondBody, cancellationToken);
        foreach (XAttribute id in second.Root.Attributes().Where(attribute => attribute.Name == "id" || attribute.Name == XNamespace.Xml + "id")) shared.Add(id.Value);
        VerifyMergeLocalReferences(second.Root, shared, secondPath, cancellationToken);
        (string Path, string? Fragment) Map(EpubReference reference) => reference.ContainerPath == secondPath ?
            (firstPath, string.IsNullOrEmpty(reference.Fragment) || shared.Contains(reference.Fragment!) ? boundaryId : reference.Fragment) :
            (reference.ContainerPath!, reference.Fragment);
        RewriteMovedXml(first, firstPath, firstPath, string.Empty, string.Empty, cancellationToken, Map, removeHtmlBase: true);
        RewriteMovedXml(second, secondPath, firstPath, string.Empty, string.Empty, cancellationToken, Map, removeHtmlBase: true);
        if (!SameMergeAttributes(first.Root, second.Root) || !SameMergeAttributes(firstBody, secondBody))
            throw new NotSupportedException("Root/body attributes resolve differently after reference repair; resolve their styling and semantics before merging.");
        XElement firstHead = first.Root.Element(Html + "head")!, secondHead = second.Root.Element(Html + "head")!;
        if (!XNode.DeepEquals(ComparableMergeHead(firstHead), ComparableMergeHead(secondHead)) ||
            !first.Nodes().Where(node => node != first.Root).Select(node => node.ToString()).SequenceEqual(second.Nodes().Where(node => node != second.Root).Select(node => node.ToString()), StringComparer.Ordinal))
            throw new NotSupportedException("Resolve conflicting chapter heads or document instructions before merging. Styles and metadata are not silently combined or discarded.");
        var boundary = new XElement(Html + "span", new XAttribute("id", boundaryId), new XAttribute("title", secondHead.Element(Html + "title")!.Value));
        JoinMergeContainers(firstBody, secondBody, boundary, cancellationToken);
        EpubContentIdentifiers.ValidateReferences(first.Root, EpubContentIdentifiers.Collect(first.Root, firstPath, true, cancellationToken), firstPath, cancellationToken);
        if (first.Descendants(Html + "map").Attributes("name").GroupBy(attribute => attribute.Value, StringComparer.Ordinal).Any(group => group.Count() > 1))
            throw new InvalidDataException("Image-map names collide in the merged chapter.");
        byte[] merged = SerializeXml(first, _maximumEntryBytes);
        XDocument package = new XDocument(_package);
        if (package.Descendants().Attributes(XNamespace.Xml + "base").Any()) throw new NotSupportedException("Chapter merging does not support XML base declarations.");
        XAttribute[] originals = Root.DescendantsAndSelf().Attributes().ToArray(), proposed = package.Root!.DescendantsAndSelf().Attributes().ToArray();
        var packageEdits = new List<(XAttribute Attribute, string Value)>();
        var references = new HashSet<XAttribute>(PackageResourceReferences(package.Root));
        for (int index = 0; index < proposed.Length; index++) {
            if (!references.Contains(proposed[index])) continue;
            string value = RewriteMovedReference(PackagePath, null, PackagePath, null, proposed[index].Value, string.Empty, string.Empty, Map);
            if (value != proposed[index].Value) { proposed[index].Value = value; packageEdits.Add((originals[index], value)); }
        }
        package.Root.Element(Opf + "manifest")!.Elements(Opf + "item").Single(item => (string?)item.Attribute("id") == secondManifestId).Remove();
        package.Root.Element(Opf + "spine")!.Elements(Opf + "itemref").Single(item => (string?)item.Attribute("idref") == secondManifestId).Remove();
        var entries = new Dictionary<string, byte[]>(_entries, StringComparer.Ordinal);
        foreach (var group in Manifest.Where(item => item.Reference.Kind == EpubReferenceKind.Container).GroupBy(RequireLocalPath, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = group.Key;
            if (path == secondPath) continue;
            string[] types = group.Select(item => item.MediaType).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (types.Length != 1) throw new NotSupportedException("Reference repair requires unambiguous resource media types.");
            byte[] original = _entries[path];
            byte[] edited = path == firstPath ? merged : RewritePublicationResource(types[0], original, path, path, string.Empty, string.Empty, cancellationToken, Map);
            if (!ReferenceEquals(original, edited)) EnsureResourceMutationAllowed(path);
            if (edited.LongLength > _maximumEntryBytes) throw new InvalidDataException("Merge reference repair exceeds the retained entry-byte limit.");
            entries[path] = edited;
        }
        entries.Remove(secondPath);
        long delta = entries.Values.Sum(value => value.LongLength) - _retainedBytes;
        EnsurePackageBudget(package, delta);
        ValidatePublication(package, entries, new List<OfficeConversionFidelityDiagnostic>(), cancellationToken, changed: true);
        EnsurePackageBudget(package, delta);
        cancellationToken.ThrowIfCancellationRequested();
        foreach (var edit in packageEdits) edit.Attribute.Value = edit.Value;
        RequireSection("manifest").Elements(Opf + "item").Single(item => (string?)item.Attribute("id") == secondManifestId).Remove();
        secondPosition.Remove();
        _entries.Clear(); foreach (var entry in entries) _entries.Add(entry.Key, entry.Value);
        RetainMergedOrigins(firstPath, secondPath);
        _retainedBytes += delta; MarkChanged();
    }
}
