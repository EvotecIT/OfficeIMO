using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Splits a reflowable XHTML chapter before an identified block, moving that block and following
    /// content into a new chapter immediately after it. Clones the head and structural containers,
    /// repairs publication references, and adds a sibling TOC entry. Rejects separated local ID references.
    /// </summary>
    public EpubManifestItem SplitChapter(string manifestId, string boundaryId, string newManifestId, string newContainerPath,
        string title, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        RequireText(title, nameof(title)); RequireText(boundaryId, nameof(boundaryId)); VerifyAvailableId(newManifestId);
        newContainerPath = VerifyContentPath(newContainerPath);
        VerifyAvailableContainerPath(newContainerPath);
        EpubManifestItem source = RequireManifestItem(manifestId);
        XElement position = RequireRestructurablePosition(source);
        string sourcePath = RequireLocalPath(source);
        XDocument first = EditableXhtml(manifestId);
        EpubContentIdentifiers.Collect(first.Root!, sourcePath, true, cancellationToken);
        XElement body = first.Root!.Element(Html + "body") ?? throw new InvalidDataException("Chapter has no XHTML body.");
        XElement boundary = RequireContentElement(first, boundaryId);
        XElement[] ancestry = boundary.Ancestors().TakeWhile(element => element != body).ToArray();
        if (!boundary.Ancestors().Contains(body) || ancestry.Length > 64 || ancestry.Any(element => !IsSplitContainer(element)) ||
            boundary.Name.Namespace != Html || !new[] { "h1", "h2", "h3", "h4", "h5", "h6", "p", "div", "section", "article", "main", "aside",
                "blockquote", "ul", "ol", "dl", "figure", "table", "pre", "hr", "address", "header", "footer" }.Contains(boundary.Name.LocalName, StringComparer.Ordinal))
            throw new NotSupportedException("Split before a complete block inside body, div, section, article or main containers, within 64 levels.");
        var split = SplitContentContainer(body, boundary, new HashSet<XElement>(ancestry), cancellationToken);
        if (!HasSplitContent(split.First) || !HasSplitContent(split.Second)) throw new InvalidOperationException("Both split chapters must retain content.");
        XDocument second = new XDocument(first);
        body.ReplaceWith(split.First);
        second.Root!.Element(Html + "body")!.ReplaceWith(split.Second);
        XElement secondTitle = second.Root.Element(Html + "head")?.Element(Html + "title") ?? throw new InvalidDataException("Chapter has no title.");
        secondTitle.Value = title;
        HashSet<string> firstIds = EpubContentIdentifiers.Collect(first.Root!, sourcePath, true, cancellationToken);
        HashSet<string> secondIds = EpubContentIdentifiers.Collect(second.Root, newContainerPath, true, cancellationToken);
        var movedIds = new HashSet<string>(secondIds.Except(firstIds, StringComparer.Ordinal), StringComparer.Ordinal);
        EpubContentIdentifiers.ValidateReferences(first.Root!, firstIds, sourcePath, cancellationToken);
        EpubContentIdentifiers.ValidateReferences(second.Root!, secondIds, newContainerPath, cancellationToken);
        (string Path, string? Fragment) Map(EpubReference reference) =>
            (reference.ContainerPath == sourcePath && reference.Fragment != null && movedIds.Contains(reference.Fragment) ? newContainerPath : reference.ContainerPath!, reference.Fragment);
        (string Path, string? Fragment) MapSecond(EpubReference reference) =>
            (reference.ContainerPath == sourcePath && reference.Fragment != null && secondIds.Contains(reference.Fragment) ? newContainerPath : reference.ContainerPath!, reference.Fragment);
        RewriteMovedXml(first, sourcePath, sourcePath, string.Empty, string.Empty, cancellationToken, Map);
        RewriteMovedXml(second, sourcePath, newContainerPath, string.Empty, string.Empty, cancellationToken, MapSecond);
        byte[] firstBytes = SerializeXml(first, _maximumEntryBytes), secondBytes = SerializeXml(second, _maximumEntryBytes);
        XElement declaration = PrepareResourceDeclaration(newManifestId, newContainerPath, source.MediaType, secondBytes, source.Properties);
        XElement newPosition = new XElement(position);
        newPosition.SetAttributeValue("idref", newManifestId);
        // The original package id and its refinements continue to identify the first reading position.
        newPosition.Attribute("id")?.Remove();
        XDocument package = new XDocument(_package);
        if (package.Descendants().Attributes(XNamespace.Xml + "base").Any()) throw new NotSupportedException("Chapter splitting does not support XML base declarations.");
        var packageEdits = new List<(XAttribute Attribute, string Value)>();
        XAttribute[] oldAttributes = Root.DescendantsAndSelf().Attributes().ToArray();
        XAttribute[] newAttributes = package.Root!.DescendantsAndSelf().Attributes().ToArray();
        var references = new HashSet<XAttribute>(PackageResourceReferences(package.Root));
        for (int index = 0; index < newAttributes.Length; index++) {
            if (!references.Contains(newAttributes[index])) continue;
            string value = RewriteMovedReference(PackagePath, null, PackagePath, null, newAttributes[index].Value, string.Empty, string.Empty, Map);
            if (value == newAttributes[index].Value) continue;
            newAttributes[index].Value = value; packageEdits.Add((oldAttributes[index], value));
        }
        package.Root.Element(Opf + "manifest")!.Elements(Opf + "item").Single(element => (string?)element.Attribute("id") == manifestId).AddAfterSelf(new XElement(declaration));
        package.Root.Element(Opf + "spine")!.Elements(Opf + "itemref").Single(element => (string?)element.Attribute("idref") == manifestId).AddAfterSelf(new XElement(newPosition));
        var entries = new Dictionary<string, byte[]>(_entries, StringComparer.Ordinal);
        string navigationPath = NavigationPath();
        foreach (var group in Manifest.Where(item => item.Reference.Kind == EpubReferenceKind.Container).GroupBy(RequireLocalPath, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = group.Key;
            string[] types = group.Select(item => item.MediaType).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (types.Length != 1) throw new NotSupportedException("Reference repair requires unambiguous resource media types.");
            byte[] original = _entries[path];
            byte[] edited = path == sourcePath ? firstBytes : (path == navigationPath || HasMediaType(types[0], "application/x-dtbncx+xml")) ?
                PrepareSplitNavigation(path, sourcePath, newContainerPath, boundaryId, title, Map, cancellationToken) :
                RewritePublicationResource(types[0], original, path, path, string.Empty, string.Empty, cancellationToken, Map);
            if (!ReferenceEquals(original, edited)) EnsureResourceMutationAllowed(path);
            if (edited.LongLength > _maximumEntryBytes) throw new InvalidDataException("Split reference repair exceeds the retained entry-byte limit.");
            entries[path] = edited;
        }
        entries.Add(newContainerPath, secondBytes);
        long delta = entries.Values.Sum(value => value.LongLength) - _retainedBytes;
        EnsurePackageBudget(package, delta);
        ValidatePublication(package, entries, new List<OfficeConversionFidelityDiagnostic>(), cancellationToken, changed: true);
        EnsurePackageBudget(package, delta);
        cancellationToken.ThrowIfCancellationRequested();
        foreach (var edit in packageEdits) edit.Attribute.Value = edit.Value;
        RequireSection("manifest").Elements(Opf + "item").Single(element => (string?)element.Attribute("id") == manifestId).AddAfterSelf(declaration);
        position.AddAfterSelf(newPosition);
        foreach (var entry in entries) _entries[entry.Key] = entry.Value;
        _retainedBytes += delta;
        MarkChanged();
        return RequireManifestItem(newManifestId);
    }
}
