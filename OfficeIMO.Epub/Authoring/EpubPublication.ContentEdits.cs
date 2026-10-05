using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // Stage every payload and the combined retention budget before committing any
    // document. This keeps cross-document semantic edits atomic without a ZIP round-trip.
    private void CommitContentEdits(IReadOnlyDictionary<string, XDocument> documents, CancellationToken token) {
        var payloads = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        long delta = 0;
        foreach (var pair in documents) {
            token.ThrowIfCancellationRequested();
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
        EnsurePackageBudget(_package, delta);
        token.ThrowIfCancellationRequested();
        foreach (var pair in payloads) _entries[pair.Key] = pair.Value;
        _retainedBytes += delta;
        MarkChanged();
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
