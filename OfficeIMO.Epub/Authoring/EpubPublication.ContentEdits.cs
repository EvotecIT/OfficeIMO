using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Atomically replaces selected XHTML/SVG documents after validating their identifiers and the
    /// resulting publication references. Use one batch when moving targets or repairing cross-document
    /// links. A rejected or cancelled batch leaves all retained content unchanged.
    /// </summary>
    public void SetContentXml(IReadOnlyDictionary<string, XDocument> documents, CancellationToken cancellationToken = default) {
        if (documents == null) throw new ArgumentNullException(nameof(documents));
        cancellationToken.ThrowIfCancellationRequested();
        if (documents.Count == 0) return;
        CommitContentEdits(documents, cancellationToken, validatePublication: true);
    }

    // Stage every payload and the combined retention budget before committing any
    // document. This keeps cross-document semantic edits atomic without a ZIP round-trip.
    private void CommitContentEdits(IReadOnlyDictionary<string, XDocument> documents, CancellationToken token, bool validatePublication = false, Action<XElement>? packageEdit = null) {
        var payloads = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        long delta = 0;
        foreach (var pair in documents) {
            token.ThrowIfCancellationRequested();
            if (pair.Value == null) throw new ArgumentException("Content documents cannot be null.", nameof(documents));
            EpubManifestItem item = RequireManifestItem(pair.Key);
            string path = RequireLocalPath(item);
            EnsureResourceMutationAllowed(path);
            if (_encryption.Any(encryption => encryption.Path == path))
                throw new NotSupportedException("Encrypted/obfuscated content cannot be edited.");
            ValidateContent(pair.Value, item.MediaType);
            HashSet<string> ids = EpubContentIdentifiers.Collect(pair.Value.Root!, path, true, token);
            EpubContentIdentifiers.ValidateReferences(pair.Value.Root!, ids, path, token);
            byte[] bytes = SerializeXml(pair.Value, _maximumEntryBytes);
            if (payloads.ContainsKey(path)) throw new InvalidDataException("Content edits select the same resource through multiple manifest declarations.");
            payloads.Add(path, bytes);
            delta += bytes.LongLength - _entries[path].LongLength;
        }
        var package = new XDocument(_package);
        packageEdit?.Invoke(package.Root!);
        EnsurePackageBudget(package, delta);
        if (validatePublication) {
            var proposed = new Dictionary<string, byte[]>(_entries, StringComparer.Ordinal);
            foreach (var pair in payloads) proposed[pair.Key] = pair.Value;
            ValidatePublication(package, proposed, new List<OfficeConversionFidelityDiagnostic>(), token, changed: true);
            EnsurePackageBudget(package, delta);
        }
        token.ThrowIfCancellationRequested();
        packageEdit?.Invoke(Root);
        foreach (var pair in payloads) _entries[pair.Key] = pair.Value;
        _retainedBytes += delta;
        MarkChanged();
    }

    private IEnumerable<XNode> ParseAuthoringFragment(string xhtml) {
        using (var reader = XmlReader.Create(new StringReader("<div xmlns='" + Html.NamespaceName + "' xmlns:epub='" + Ops.NamespaceName + "'>" + xhtml + "</div>"),
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = _maximumEntryBytes }))
            return XElement.Load(reader, LoadOptions.PreserveWhitespace).Nodes().ToArray();
    }

    private XDocument EditableXhtml(string manifestId) {
        EpubManifestItem item = RequireManifestItem(manifestId);
        if (!HasMediaType(item.MediaType, "application/xhtml+xml")) throw new NotSupportedException("Semantic authoring requires XHTML content.");
        string path = RequireLocalPath(item);
        if (_encryption.Any(encryption => encryption.Path == path)) throw new NotSupportedException("Encrypted content cannot be edited.");
        XDocument document = GetContentXml(manifestId);
        ValidateContent(document, item.MediaType);
        return document;
    }

    private static XElement RequireContentElement(XDocument document, string id) {
        XElement[] matches = document.Root!.DescendantsAndSelf().Where(element =>
            (string?)element.Attribute("id") == id || (string?)element.Attribute(XNamespace.Xml + "id") == id).ToArray();
        if (matches.Length != 1) throw new InvalidDataException("Expected one content element with identifier: " + id);
        return matches[0];
    }
}
