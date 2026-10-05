using System.Globalization;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Atomically replaces a single chapter's sequential narration and recalculates duration metadata.
    /// Preserves cue ids for surviving text targets, manifest identity and metadata attributes. Rejects
    /// shared or encrypted overlays, richer SMIL structures, and referenced cues that would be removed.
    /// Audio resources remain unchanged; their durations are caller declarations.
    /// </summary>
    public void ReplaceMediaOverlay(string overlayManifestId, EpubMediaOverlay overlay, CancellationToken cancellationToken = default) {
        if (overlay == null) throw new ArgumentNullException(nameof(overlay));
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Media-overlay authoring requires EPUB 3.");
        if (EpubVocabulary.Expand(Root, "media:duration") != OverlayVocabulary + "duration")
            throw new InvalidDataException("The media prefix is mapped to a different vocabulary.");
        EpubManifestItem item = RequireManifestItem(overlayManifestId);
        if (!HasMediaType(item.MediaType, "application/smil+xml")) throw new ArgumentException("Expected a SMIL manifest resource.", nameof(overlayManifestId));
        string path = RequireLocalPath(item);
        EnsureResourceMutationAllowed(path);
        if (Manifest.Count(resource => resource.Reference.ContainerPath == path) != 1 || _encryption.Any(entry => entry.Path == path))
            throw new NotSupportedException("Overlay replacement requires a uniquely declared, unencrypted resource.");
        EnsureRemovalPreservesRootfiles(path);
        EpubManifestItem[] owners = Manifest.Where(resource => resource.MediaOverlayId == overlayManifestId).ToArray();
        if (owners.Length != 1) throw new NotSupportedException("Overlay replacement requires exactly one associated content document.");
        string contentPath = RequireLocalPath(owners[0]);
        XDocument prior = GetContentXml(overlayManifestId);
        Dictionary<string, string?> priorCueIds = ReadReplaceableOverlay(prior, path, contentPath, cancellationToken);
        var prepared = PrepareMediaOverlay(owners[0], path, overlay, cancellationToken);
        var reservedIds = new HashSet<string>(priorCueIds.Values.Where(id => id != null).Select(id => id!), StringComparer.Ordinal);
        var retainedIds = new HashSet<string>(StringComparer.Ordinal);
        int nextId = 0;
        foreach (XElement cue in prepared.Document.Descendants(Smil + "par")) {
            cancellationToken.ThrowIfCancellationRequested();
            string target = EpubReference.Resolve(path, (string)cue.Element(Smil + "text")!.Attribute("src")!).Fragment!;
            if (!priorCueIds.TryGetValue(target, out string? id) || id == null) {
                do { id = "cue" + (nextId++).ToString(CultureInfo.InvariantCulture); } while (reservedIds.Contains(id));
                reservedIds.Add(id);
            }
            cue.SetAttributeValue("id", id); retainedIds.Add(id);
        }
        var removedIds = new HashSet<string>(priorCueIds.Values.Where(id => id != null && !retainedIds.Contains(id)).Select(id => id!), StringComparer.Ordinal);
        EnsureOverlayCueRemovalSafe(path, removedIds, cancellationToken);
        byte[] payload = SerializeXml(prepared.Document, _maximumEntryBytes);
        long delta = payload.LongLength - _entries[path].LongLength;
        XElement[] oldDurations = OverlayDurations(Root).Where(meta => ReferencesPackageId((string?)meta.Attribute("refines") ?? string.Empty, overlayManifestId)).ToArray();
        if (oldDurations.Length > 1) throw new InvalidDataException("The overlay has ambiguous duration metadata.");
        TimeSpan total = MediaOverlayTotal(prepared.Duration, overlayManifestId);
        EditPackageElement(Root, proposed => {
            cancellationToken.ThrowIfCancellationRequested();
            XElement metadata = proposed.Element(Opf + "metadata")!;
            XElement? duration = OverlayDurations(proposed).SingleOrDefault(meta => ReferencesPackageId((string?)meta.Attribute("refines") ?? string.Empty, overlayManifestId));
            if (duration == null) metadata.Add(DurationMetadata(prepared.Duration, "#" + overlayManifestId));
            else duration.Value = EpubSmilClock.Format(prepared.Duration);
            XElement? totalMetadata = OverlayDurations(proposed).SingleOrDefault(meta => meta.Attribute("refines") == null);
            if (totalMetadata == null) metadata.Add(DurationMetadata(total, null));
            else totalMetadata.Value = EpubSmilClock.Format(total);
        }, delta);
        _entries[path] = payload;
        _retainedBytes += delta;
        MarkChanged();
    }

    private Dictionary<string, string?> ReadReplaceableOverlay(XDocument document, string path, string contentPath, CancellationToken token) {
        if (document.DescendantNodes().Any(node => node is XProcessingInstruction || node is XDocumentType))
            throw new NotSupportedException("Overlay replacement cannot preserve processing instructions or document types.");
        EpubContentIdentifiers.Collect(document.Root ?? throw new InvalidDataException("Overlay has no root."), path, true, token);
        XElement root = document.Root;
        RequireOverlayShape(root, Smil + "smil", new[] { XName.Get("version") }, 1);
        if ((string?)root.Attribute("version") != "3.0") throw new NotSupportedException("Expected SMIL 3.0.");
        XElement body = root.Elements().Single();
        RequireOverlayShape(body, Smil + "body", Array.Empty<XName>(), 1);
        XElement sequence = body.Elements().Single();
        RequireOverlayShape(sequence, Smil + "seq", new[] { Ops + "textref" }, null);
        EpubReference sequenceTarget = EpubReference.Resolve(path, (string?)sequence.Attribute(Ops + "textref") ?? string.Empty);
        if (sequenceTarget.Kind != EpubReferenceKind.Container || sequenceTarget.ContainerPath != contentPath || sequenceTarget.Fragment != null)
            throw new NotSupportedException("The overlay sequence must target its associated whole XHTML document.");
        var result = new Dictionary<string, string?>(StringComparer.Ordinal);
        foreach (XElement cue in sequence.Elements()) {
            token.ThrowIfCancellationRequested();
            if (result.Count >= 10000) throw new NotSupportedException("Overlay replacement supports at most 10,000 cues.");
            RequireOverlayShape(cue, Smil + "par", new[] { XName.Get("id") }, 2);
            XElement[] children = cue.Elements().ToArray();
            XElement text = children.SingleOrDefault(element => element.Name == Smil + "text") ?? throw new NotSupportedException("Cue requires one text target.");
            XElement audio = children.SingleOrDefault(element => element.Name == Smil + "audio") ?? throw new NotSupportedException("Cue requires one audio clip.");
            RequireOverlayShape(text, Smil + "text", new[] { XName.Get("src") }, 0);
            RequireOverlayShape(audio, Smil + "audio", new[] { XName.Get("src"), XName.Get("clipBegin"), XName.Get("clipEnd") }, 0);
            EpubReference target = EpubReference.Resolve(path, (string?)text.Attribute("src") ?? string.Empty);
            if (target.Kind != EpubReferenceKind.Container || target.ContainerPath != contentPath || string.IsNullOrEmpty(target.Fragment) || result.ContainsKey(target.Fragment!))
                throw new NotSupportedException("Replacement requires distinct text fragments in the associated chapter.");
            string? id = (string?)cue.Attribute("id");
            if (id != null) XmlConvert.VerifyNCName(id);
            result.Add(target.Fragment!, id);
        }
        if (result.Count == 0) throw new NotSupportedException("The overlay has no replaceable cues.");
        return result;
    }

    private static void RequireOverlayShape(XElement element, XName name, XName[] attributes, int? children) {
        if (element.Name != name || element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && !attributes.Contains(attribute.Name)) ||
            children.HasValue && element.Elements().Count() != children.Value ||
            element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            throw new NotSupportedException("Replacement supports the single-sequence text/audio profile without additional SMIL structures or attributes.");
    }

    private void EnsureOverlayCueRemovalSafe(string path, HashSet<string> removedIds, CancellationToken token) {
        if (removedIds.Count == 0) return;
        (string Path, string? Fragment) Check(EpubReference reference) {
            if (reference.ContainerPath == path && reference.Fragment != null && removedIds.Contains(reference.Fragment))
                throw new InvalidOperationException("A removed narration cue is still referenced: " + reference.Fragment);
            return (reference.ContainerPath!, reference.Fragment);
        }
        if (Root.DescendantsAndSelf().Attributes(XNamespace.Xml + "base").Any())
            throw new NotSupportedException("Cue removal cannot inspect package XML base declarations.");
        foreach (XAttribute reference in PackageResourceReferences(Root)) {
            token.ThrowIfCancellationRequested();
            Check(EpubReference.Resolve(PackagePath, reference.Value));
        }
        foreach (var group in Manifest.Where(resource => resource.Reference.Kind == EpubReferenceKind.Container).GroupBy(RequireLocalPath, StringComparer.Ordinal)) {
            token.ThrowIfCancellationRequested();
            if (group.Key == path) continue;
            string[] types = group.Select(resource => resource.MediaType).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (types.Length != 1 || _encryption.Any(entry => entry.Path == group.Key && entry.RequiresDecryption))
                throw new NotSupportedException("Cue removal requires inspectable resources with unambiguous media types.");
            // The canonical reference walker runs with an identity map; no payload changes are committed.
            RewritePublicationResource(types[0], _entries[group.Key], group.Key, group.Key, string.Empty, string.Empty, token, Check);
        }
    }
}
