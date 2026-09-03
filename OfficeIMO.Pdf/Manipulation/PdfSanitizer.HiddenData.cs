namespace OfficeIMO.Pdf;

internal static partial class PdfSanitizer {
    private static readonly HashSet<string> TechnicalInfoKeys = new HashSet<string>(StringComparer.Ordinal) {
        "Producer", "CreationDate", "ModDate", "Trapped"
    };

    private static PdfSanitizationReport BuildReport(
        byte[] pdf,
        PdfSanitizationOptions policy,
        PdfLoadOptions? readOptions,
        IReadOnlyList<PdfSanitizationFinding> findings) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        var parsed = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, policy.CancellationToken);
        PdfDocumentSecurityInfo baseline = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            readOptions,
            includeParsedDetails: false,
            cancellationToken: policy.CancellationToken);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            parsed.Map,
            parsed.TrailerRaw,
            baseline,
            readOptions,
            policy.CancellationToken);
        int userMetadata = policy.ShouldRemoveUserMetadata ? CountUserMetadataEntries(parsed.Map, security) : 0;
        int embeddedFiles = policy.ShouldRemoveEmbeddedFiles
            ? PdfAttachmentExtractor.InspectAttachments(
                parsed.Map,
                parsed.TrailerRaw,
                readOptions?.Limits,
                policy.CancellationToken).Count
            : 0;
        int commentsAndMarkup = policy.ShouldRemoveCommentsAndMarkup
            ? CountSelectedCommentAnnotations(parsed.Map, policy)
            : 0;
        int bookmarks = policy.ShouldRemoveBookmarks
            ? CountOutlineItems(parsed.Map, security, policy.CancellationToken)
            : 0;
        int optionalContent = policy.ShouldRemoveOptionalContent
            ? CountOptionalContentGroups(parsed.Map, security)
            : 0;
        return new PdfSanitizationReport(
            findings,
            userMetadata,
            embeddedFiles,
            commentsAndMarkup,
            bookmarks,
            optionalContent);
    }

    private static int CountUserMetadataEntries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security) {
        int count = 0;
        if (security.InfoObjectNumber.HasValue &&
            objects.TryGetValue(security.InfoObjectNumber.Value, out PdfIndirectObject? infoObject) &&
            infoObject.Value is PdfDictionary info) {
            foreach (string key in info.Items.Keys) {
                if (!TechnicalInfoKeys.Contains(key)) count++;
            }
        }
        if (security.RootObjectNumber.HasValue &&
            objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) &&
            rootObject.Value is PdfDictionary catalog &&
            catalog.Items.ContainsKey("Metadata")) {
            count++;
        }
        return count;
    }

    private static int CountSelectedCommentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfSanitizationOptions policy) {
        int count = 0;
        var visited = new HashSet<PdfDictionary>();
        var counted = new HashSet<PdfDictionary>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            CountSelectedCommentAnnotations(objects, indirect.Value, policy, visited, counted, ref count);
        }
        return count;
    }

    private static void CountSelectedCommentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfSanitizationOptions policy,
        HashSet<PdfDictionary> visited,
        HashSet<PdfDictionary> counted,
        ref int count) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        if (value is PdfStream stream) value = stream.Dictionary;
        if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                if (array.Items[i] is not PdfReference) {
                    CountSelectedCommentAnnotations(objects, array.Items[i], policy, visited, counted, ref count);
                }
            }
            return;
        }
        if (value is not PdfDictionary dictionary || !visited.Add(dictionary)) return;
        if (Resolve(objects, dictionary.Get<PdfObject>("Type")) is PdfName type &&
            string.Equals(type.Name, "Annot", StringComparison.Ordinal) &&
            IsSelectedCommentAnnotation(objects, dictionary, policy) &&
            counted.Add(dictionary)) {
            count = checked(count + 1);
        }
        if (dictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) &&
            Resolve(objects, annotationsObject) is PdfArray annotations) {
            for (int i = 0; i < annotations.Items.Count; i++) {
                if (Resolve(objects, annotations.Items[i]) is PdfDictionary annotation &&
                    IsSelectedCommentAnnotation(objects, annotation, policy) &&
                    counted.Add(annotation)) {
                    count = checked(count + 1);
                }
            }
        }
        foreach (PdfObject child in dictionary.Items.Values) {
            if (child is not PdfReference) {
                CountSelectedCommentAnnotations(objects, child, policy, visited, counted, ref count);
            }
        }
    }

    private static int CountOutlineItems(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        System.Threading.CancellationToken cancellationToken) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("Outlines", out PdfObject? outlinesObject) ||
            Resolve(objects, outlinesObject) is not PdfDictionary outlines) {
            return 0;
        }
        bool hasFirst = outlines.Items.TryGetValue("First", out PdfObject? first);
        bool hasLast = outlines.Items.TryGetValue("Last", out PdfObject? last);
        if (!hasFirst && !hasLast) return 1;

        int count = 0;
        var pending = new Stack<PdfObject>();
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var visitedDictionaries = new HashSet<PdfDictionary>();
        if (first is not null) pending.Push(first);
        if (last is not null) pending.Push(last);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject current = pending.Pop();
            if (current is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) continue;
                current = indirect.Value;
            }
            if (current is not PdfDictionary item || !visitedDictionaries.Add(item)) continue;
            count = checked(count + 1);
            PushOutlineLinks(item, pending);
        }
        return count == 0 ? 1 : count;
    }

    private static int CountOptionalContentGroups(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContentObject) ||
            Resolve(objects, optionalContentObject) is not PdfDictionary optionalContent) {
            return 0;
        }

        if (!optionalContent.Items.TryGetValue("OCGs", out PdfObject? groupsObject) ||
            Resolve(objects, groupsObject) is not PdfArray groups) {
            return 1;
        }

        var counted = new HashSet<PdfDictionary>();
        for (int index = 0; index < groups.Items.Count; index++) {
            if (Resolve(objects, groups.Items[index]) is PdfDictionary group) counted.Add(group);
        }
        return Math.Max(1, counted.Count);
    }

    private static void SanitizeDocumentContainers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        PdfSanitizationOptions policy) {
        if (policy.ShouldRemoveUserMetadata && security.InfoObjectNumber.HasValue &&
            objects.TryGetValue(security.InfoObjectNumber.Value, out PdfIndirectObject? infoObject) &&
            infoObject.Value is PdfDictionary info) {
            foreach (string key in info.Items.Keys.Where(static key => !TechnicalInfoKeys.Contains(key)).ToArray()) {
                info.Items.Remove(key);
            }
        }

        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("The active PDF trailer does not reference a readable catalog object.");
        }

        if (policy.ShouldRemoveUserMetadata && catalog.Items.TryGetValue("Metadata", out PdfObject? metadata)) {
            RemoveSelectedMetadataAliases(objects, metadata);
            NeutralizeReferencedStream(objects, metadata);
            catalog.Items.Remove("Metadata");
        }
        if (policy.ShouldRemoveBookmarks) {
            if (catalog.Items.TryGetValue("Outlines", out PdfObject? outlines)) ClearOutlineTree(objects, outlines, policy.CancellationToken);
            catalog.Items.Remove("Outlines");
            if (HasCatalogPageMode(objects, catalog, "UseOutlines")) catalog.Items.Remove("PageMode");
        }
        if (policy.ShouldRemoveEmbeddedFiles && HasCatalogPageMode(objects, catalog, "UseAttachments")) {
            catalog.Items.Remove("PageMode");
        }
        if (policy.ShouldRemoveOptionalContent) {
            if (catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContent)) {
                NeutralizeObjectGraph(objects, optionalContent, policy.CancellationToken);
            }
            catalog.Items.Remove("OCProperties");
            if (HasCatalogPageMode(objects, catalog, "UseOC")) catalog.Items.Remove("PageMode");
        }
    }

    private static bool HasCatalogPageMode(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        string expected) =>
        Resolve(objects, catalog.Get<PdfObject>("PageMode")) is PdfName pageMode &&
        string.Equals(pageMode.Name, expected, StringComparison.Ordinal);

    private static bool ShouldRemoveAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        PdfSanitizationOptions policy) {
        if (Resolve(objects, annotation.Get<PdfObject>("Subtype")) is not PdfName subtype) return false;
        if (string.Equals(subtype.Name, "FileAttachment", StringComparison.Ordinal) && policy.ShouldRemoveEmbeddedFiles) {
            return policy.ContentKindsToRemove.HasValue || policy.RemoveRichMedia;
        }
        return policy.ShouldRemoveCommentAnnotation(subtype.Name) || policy.ShouldRemoveLegacyRichAnnotation(subtype.Name);
    }

    private static bool IsSelectedCommentAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        PdfSanitizationOptions policy) {
        if (Resolve(objects, annotation.Get<PdfObject>("Subtype")) is not PdfName subtype) return false;
        return policy.ShouldRemoveCommentAnnotation(subtype.Name) || policy.ShouldRemoveLegacyRichAnnotation(subtype.Name);
    }

    private static void RemoveOptionalContentAssociation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary) {
        string? type = Resolve(objects, dictionary.Get<PdfObject>("Type")) is PdfName typeName ? typeName.Name : null;
        if (string.Equals(type, "OCG", StringComparison.Ordinal) || string.Equals(type, "OCMD", StringComparison.Ordinal)) {
            dictionary.Items.Clear();
            return;
        }
        dictionary.Items.Remove("OC");
        dictionary.Items.Remove("OCGs");
    }

    private static void NeutralizeReferencedStream(Dictionary<int, PdfIndirectObject> objects, PdfObject value) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        while (value is PdfReference reference && visited.Add((reference.ObjectNumber, reference.Generation)) &&
               PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
            if (indirect.Value is PdfStream) {
                objects[reference.ObjectNumber] = new PdfIndirectObject(
                    indirect.ObjectNumber,
                    indirect.Generation,
                    new PdfStream(new PdfDictionary(), Array.Empty<byte>()));
                return;
            }
            value = indirect.Value;
        }
    }

    private static void ClearOutlineTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject outlines,
        System.Threading.CancellationToken cancellationToken) {
        var pending = new Stack<PdfObject>();
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var visitedDictionaries = new HashSet<PdfDictionary>();
        pending.Push(outlines);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject current = pending.Pop();
            if (current is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) continue;
                current = indirect.Value;
            }
            if (current is not PdfDictionary dictionary || !visitedDictionaries.Add(dictionary)) continue;
            PushOutlineLinks(dictionary, pending);
            dictionary.Items.Clear();
        }
    }

    private static void RemoveSelectedMetadataAliases(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject selectedMetadata) {
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value is PdfStream stream
                ? stream.Dictionary
                : indirect.Value as PdfDictionary;
            if (dictionary is null ||
                !dictionary.Items.TryGetValue("Metadata", out PdfObject? candidate)) continue;
            if (ReferencesSameObject(objects, candidate, selectedMetadata)) dictionary.Items.Remove("Metadata");
        }
    }

    private static bool ReferencesSameObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject left,
        PdfObject right) {
        if (left is PdfReference leftReference && right is PdfReference rightReference) {
            return leftReference.ObjectNumber == rightReference.ObjectNumber &&
                leftReference.Generation == rightReference.Generation;
        }
        return ReferenceEquals(Resolve(objects, left), Resolve(objects, right));
    }

    private static void NeutralizeObjectGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject root,
        System.Threading.CancellationToken cancellationToken) {
        var pending = new Stack<PdfObject>();
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var visitedObjects = new HashSet<PdfObject>();
        pending.Push(root);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject current = pending.Pop();
            if (current is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) continue;
                current = indirect.Value;
            }
            if (!visitedObjects.Add(current)) continue;
            if (current is PdfStream stream) {
                foreach (PdfObject child in stream.Dictionary.Items.Values) pending.Push(child);
                stream.Dictionary.Items.Clear();
                continue;
            }
            if (current is PdfArray array) {
                for (int index = 0; index < array.Items.Count; index++) pending.Push(array.Items[index]);
                array.Items.Clear();
                continue;
            }
            if (current is not PdfDictionary dictionary) continue;
            foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
            dictionary.Items.Clear();
        }
    }

    private static void PushOutlineLinks(PdfDictionary dictionary, Stack<PdfObject> pending) {
        foreach (string key in new[] { "First", "Last", "Next", "Prev" }) {
            if (dictionary.Items.TryGetValue(key, out PdfObject? linked)) pending.Push(linked);
        }
    }
}
