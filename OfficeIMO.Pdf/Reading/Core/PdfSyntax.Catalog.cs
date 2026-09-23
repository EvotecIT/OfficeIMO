using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static void ThrowIfEncryptedXrefStream(Dictionary<int, PdfIndirectObject> map) {
        foreach (var entry in map.Values) {
            PdfDictionary? dictionary = entry.Value switch {
                PdfDictionary directDictionary => directDictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };

            if (dictionary is not null && dictionary.Items.ContainsKey("Encrypt")) {
                throw new NotSupportedException("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
            }
        }
    }

    internal static PdfDictionary? FindCatalog(Dictionary<int, PdfIndirectObject> map, string? trailerRaw = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (TryGetTrailerRootReference(trailerRaw, out PdfReference rootReference) &&
            PdfObjectLookup.TryGet(map, rootReference, out var rootObject) &&
            rootObject.Value is PdfDictionary rootDictionary &&
            rootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return rootDictionary;
        }

        if (TryGetXrefStreamRootReference(map, out rootReference, cancellationToken) &&
            PdfObjectLookup.TryGet(map, rootReference, out rootObject) &&
            rootObject.Value is PdfDictionary xrefRootDictionary &&
            xrefRootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return xrefRootDictionary;
        }

        return FindCatalogByScan(map, cancellationToken);
    }

    private static PdfDictionary? FindCatalogByScan(Dictionary<int, PdfIndirectObject> map, CancellationToken cancellationToken) {
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Value is PdfDictionary dictionary &&
                dictionary.Get<PdfName>("Type")?.Name == "Catalog") {
                return dictionary;
            }
        }

        return null;
    }

    private static bool TryGetTrailerRootReference(string? trailerRaw, out PdfReference reference) {
        return TryGetTrailerReference(trailerRaw, "Root", limits: null, out reference);
    }

    private static bool TryGetXrefStreamRootReference(Dictionary<int, PdfIndirectObject> map, out PdfReference reference, CancellationToken cancellationToken) {
        reference = null!;
        int highestMatchingObjectNumber = int.MinValue;
        foreach (var entry in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.ObjectNumber <= highestMatchingObjectNumber) {
                continue;
            }

            PdfDictionary? dictionary = entry.Value switch {
                PdfStream stream => stream.Dictionary,
                PdfDictionary directDictionary => directDictionary,
                _ => null
            };

            if (dictionary?.Get<PdfName>("Type")?.Name == "XRef" &&
                dictionary.Items.TryGetValue("Root", out var root) &&
                root is PdfReference rootReference &&
                ResolveObject(map, rootReference) is PdfDictionary rootDictionary &&
                rootDictionary.Get<PdfName>("Type")?.Name == "Catalog" &&
                rootReference.Generation >= 0) {
                reference = rootReference;
                highestMatchingObjectNumber = entry.ObjectNumber;
            }
        }

        return highestMatchingObjectNumber != int.MinValue;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> map, PdfObject? value) {
        return PdfObjectLookup.Resolve(map, value);
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfArray destination) {
        return destination.Items.Count > 0 &&
            destination.Items[0] is PdfReference pageReference &&
            PdfObjectLookup.TryGet(map, pageReference, out var pageObject) &&
            pageObject.Value is PdfDictionary pageDictionary &&
            pageDictionary.Get<PdfName>("Type")?.Name == "Page";
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfObject destination) {
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        while (true) {
            if (destination is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                destination = indirect.Value;
                continue;
            }

            if (destination is PdfDictionary dictionary &&
                dictionary.Items.TryGetValue("D", out var explicitDestination)) {
                destination = explicitDestination;
                continue;
            }

            return destination is PdfArray array && IsDestinationForKnownPage(map, array);
        }
    }

    private static bool IsSupportedGoToActionDictionary(Dictionary<int, PdfIndirectObject> map, PdfDictionary dictionary) {
        return dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            IsDestinationForKnownPage(map, destination);
    }

    private static bool IsSupportedOutlineAction(Dictionary<int, PdfIndirectObject> map, PdfObject action) {
        return ResolveObject(map, action) is PdfDictionary dictionary &&
            IsSupportedGoToActionDictionary(map, dictionary);
    }

    private static bool IsSimpleCatalogDictionary(PdfDictionary dictionary) {
        foreach (var value in dictionary.Items.Values) {
            if (!IsSimpleCatalogValue(value)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSimpleCatalogValue(PdfObject value) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSimpleCatalogValue(item)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsSupportedCatalogXmpMetadataStream(Dictionary<int, PdfIndirectObject> map, PdfObject value) {
        if (value is not PdfReference reference ||
            !PdfObjectLookup.TryGet(map, reference, out var indirect) ||
            indirect.Value is not PdfStream stream ||
            stream.Dictionary.Get<PdfName>("Type")?.Name != "Metadata" ||
            stream.Dictionary.Get<PdfName>("Subtype")?.Name != "XML") {
            return false;
        }

        foreach (var entry in stream.Dictionary.Items) {
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                continue;
            }

            if (!IsSimpleCatalogValue(entry.Value)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSupportedCatalogMetadataGraph(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }

                if (!PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                return !IsPageDictionary(indirect.Value) &&
                    IsSupportedCatalogMetadataGraph(map, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }

                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfStream stream:
                if (IsPageDictionary(stream.Dictionary)) {
                    return false;
                }

                foreach (var item in stream.Dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsSupportedOutlineGraph(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }

                if (!PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                return IsPageDictionary(indirect.Value) ||
                    IsSupportedOutlineGraph(map, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedOutlineGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return true;
                }

                if (dictionary.Items.ContainsKey("AA")) {
                    return false;
                }

                if (dictionary.Items.TryGetValue("A", out var action) &&
                    !IsSupportedOutlineAction(map, action)) {
                    return false;
                }

                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedOutlineGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsPageDictionary(PdfObject value) {
        return value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page";
    }

    private static bool TryGetEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject names,
        out PdfObject embeddedFiles) {
        embeddedFiles = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveObject(map, names) as PdfDictionary;
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("EmbeddedFiles", out var embeddedFileTree)) {
            return false;
        }

        embeddedFiles = embeddedFileTree;
        return true;
    }

    private static bool TryGetNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject names,
        out PdfObject namedDestinations) {
        namedDestinations = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveObject(map, names) as PdfDictionary;
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("Dests", out var namedDestinationTree)) {
            return false;
        }

        namedDestinations = namedDestinationTree;
        return true;
    }

    private static bool IsSupportedPageLabelTree(Dictionary<int, PdfIndirectObject> map, PdfObject pageLabels) {
        PdfDictionary? tree = ResolveObject(map, pageLabels) as PdfDictionary;
        if (tree is null ||
            tree.Items.ContainsKey("Kids") ||
            !tree.Items.TryGetValue("Nums", out var numsObject) ||
            ResolveObject(map, numsObject) is not PdfArray nums ||
            nums.Items.Count % 2 != 0) {
            return false;
        }

        for (int i = 0; i < nums.Items.Count; i += 2) {
            if (ResolveObject(map, nums.Items[i]) is not PdfNumber pageIndex ||
                pageIndex.Value < 0 ||
                pageIndex.Value > int.MaxValue ||
                Math.Truncate(pageIndex.Value) != pageIndex.Value ||
                ResolveObject(map, nums.Items[i + 1]) is not PdfDictionary labelDictionary) {
                return false;
            }

            foreach (var value in labelDictionary.Items.Values) {
                if (!IsSimpleCatalogValue(value)) {
                    return false;
                }
            }
        }

        return true;
    }

    private static bool IsSupportedNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject namedDestinations,
        PdfReadLimits limits) {
        int traversedNodes = 0;
        return TryCollectNamedDestinationNameTreeEntries(map, namedDestinations, new HashSet<int>(), 0, limits, ref traversedNodes);
    }

    private static bool TryCollectNamedDestinationNameTreeEntries(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences,
        int depth,
        PdfReadLimits limits,
        ref int traversedNodes) {
        if (depth > limits.MaxNameTreeDepth) {
            return false;
        }

        if (value is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                return false;
            }

            if (++traversedNodes > limits.MaxNameTreeNodes) {
                return false;
            }

            value = indirect.Value;
        }

        if (value is not PdfDictionary tree) {
            return false;
        }

        bool hasNames = tree.Items.TryGetValue("Names", out var namesObject);
        bool hasKids = tree.Items.TryGetValue("Kids", out var kidsObject);
        if (hasNames && hasKids) {
            return false;
        }

        if (hasNames) {
            if (ResolveObject(map, namesObject) is not PdfArray names ||
                names.Items.Count % 2 != 0) {
                return false;
            }

            for (int i = 0; i < names.Items.Count; i += 2) {
                if (names.Items[i] is not PdfStringObj) {
                    return false;
                }

                PdfObject? destination = ResolveObject(map, names.Items[i + 1]);
                if (destination is null || !IsDestinationForKnownPage(map, destination)) {
                    return false;
                }
            }
        }

        if (hasKids) {
            if (ResolveObject(map, kidsObject) is not PdfArray kids) {
                return false;
            }

            foreach (var kid in kids.Items) {
                if (kid is not PdfReference) {
                    return false;
                }

                if (!TryCollectNamedDestinationNameTreeEntries(map, kid, visitedReferences, depth + 1, limits, ref traversedNodes)) {
                    return false;
                }
            }
        }

        return hasNames || hasKids;
    }

    private static bool ContainsPdfName(string text, string name) {
        if (string.IsNullOrEmpty(text)) return false;

        string token = "/" + name;
        int index = 0;
        while (index < text.Length) {
            index = text.IndexOf(token, index, StringComparison.Ordinal);
            if (index < 0) return false;

            int after = index + token.Length;
            if (after >= text.Length || IsPdfDelimiter(text[after]) || char.IsWhiteSpace(text[after])) {
                return true;
            }

            index = after;
        }

        return false;
    }

    internal static bool ContainsAnyPdfName(string text, params string[] names) {
        for (int i = 0; i < names.Length; i++) {
            if (ContainsPdfName(text, names[i])) {
                return true;
            }
        }

        return false;
    }

    internal static bool ContainsAnyPdfName(
        string text,
        CancellationToken cancellationToken,
        params string[] names) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return ContainsAnyPdfName(text, names);
        for (int i = 0; i < names.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ContainsPdfName(text, names[i], cancellationToken)) return true;
        }
        return false;
    }

    private static bool ContainsPdfName(
        string text,
        string name,
        CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(text)) return false;
        string token = "/" + name;
        int maximumStart = text.Length - token.Length;
        for (int index = 0; index <= maximumStart; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (text[index] != '/' ||
                string.CompareOrdinal(text, index, token, 0, token.Length) != 0) {
                continue;
            }

            int after = index + token.Length;
            if (after >= text.Length || IsPdfDelimiter(text[after]) || char.IsWhiteSpace(text[after])) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }

    private static bool ContainsParsedOrFallbackPdfName(byte[] pdf, params string[] names) =>
        ContainsParsedOrFallbackPdfName(pdf, null, names);

    private static bool ContainsParsedOrFallbackPdfName(byte[] pdf, PdfLoadOptions? options, params string[] names) {
        try {
            var (objects, trailer) = ParseObjects(pdf, options, out PdfRepairReport repairReport);
            if (FindCatalog(objects, trailer) is null)
                return ContainsAnyPdfName(PdfEncoding.Latin1GetString(pdf), names);
            return ContainsAnyDocumentPdfName(pdf, objects, repairReport, names);
        } catch (Exception ex) when (ShouldSuppressParsedPdfNameException(ex, options)) {
            // Malformed or unauthenticated input retains the conservative raw probe.
            return ContainsAnyPdfName(PdfEncoding.Latin1GetString(pdf), names);
        }
    }

    internal static bool ContainsAnyDocumentPdfName(byte[] pdf, IReadOnlyDictionary<int, PdfIndirectObject> objects,
        PdfRepairReport repairReport, params string[] names) =>
        ContainsAnyParsedPdfName(objects, names) ||
        (repairReport.HasIncompleteObjectCoverage && ContainsAnyPdfName(PdfEncoding.Latin1GetString(pdf), names));

    internal static bool ContainsAnyDocumentPdfName(byte[] pdf, IReadOnlyDictionary<int, PdfIndirectObject> objects,
        PdfRepairReport repairReport, CancellationToken cancellationToken, params string[] names) {
        var requested = new HashSet<string>(StringComparer.Ordinal);
        foreach (string name in names) {
            cancellationToken.ThrowIfCancellationRequested();
            requested.Add(name);
        }
        if (CollectParsedPdfNames(objects, requested, cancellationToken).Count > 0) return true;
        return repairReport.HasIncompleteObjectCoverage &&
            ContainsAnyPdfName(PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken), cancellationToken, names);
    }

    internal static bool ContainsAnyParsedPdfName(
        IReadOnlyDictionary<int, PdfIndirectObject> objects,
        params string[] names) {
        var nameSet = new HashSet<string>(names, StringComparer.Ordinal);
        Func<string, bool> contains = nameSet.Contains;
        foreach (PdfIndirectObject indirectObject in objects.Values) {
            if (VisitParsedPdfNames(indirectObject.Value, contains)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Collects only requested names from the parsed object graph in one pass. String values and
    /// stream payloads are intentionally excluded; callers handle incomplete coverage separately.
    /// </summary>
    internal static HashSet<string> CollectParsedPdfNames(
        IReadOnlyDictionary<int, PdfIndirectObject> objects,
        HashSet<string> requestedNames,
        CancellationToken cancellationToken) {
        var found = new HashSet<string>(StringComparer.Ordinal);
        bool Collect(string name) {
            if (requestedNames.Contains(name)) found.Add(name);
            return false;
        }
        Func<string, bool> collect = Collect;

        foreach (PdfIndirectObject indirectObject in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            VisitParsedPdfNames(indirectObject.Value, collect, cancellationToken);
        }

        return found;
    }

    /// <summary>Matches requested name groups reachable from one parsed root in a single bounded graph walk.</summary>
    internal static int MatchReachableParsedPdfNameGroups(
        PdfObject root,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<HashSet<string>> requestedNameGroups,
        CancellationToken cancellationToken) {
        if (requestedNameGroups.Count == 0 || requestedNameGroups.Count > 30) {
            throw new ArgumentOutOfRangeException(
                nameof(requestedNameGroups),
                requestedNameGroups.Count,
                "Reachable PDF name matching requires between 1 and 30 groups.");
        }

        int matchedGroups = 0;
        int allGroups = (1 << requestedNameGroups.Count) - 1;
        var visitedObjectNumbers = new HashSet<int>();
        var visitedContainers = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>();
        pending.Push(root);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = pending.Pop();

            if (value is PdfReference reference) {
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    indirect == null ||
                    !visitedObjectNumbers.Add(reference.ObjectNumber)) {
                    continue;
                }
                pending.Push(indirect.Value);
                continue;
            }
            if (value is PdfName name) {
                if (MatchName(name.Name)) return matchedGroups;
                continue;
            }
            if (value is PdfStream stream) {
                if (visitedContainers.Add(stream)) pending.Push(stream.Dictionary);
                continue;
            }
            if (value is PdfArray array) {
                if (!visitedContainers.Add(array)) continue;
                for (int index = array.Items.Count - 1; index >= 0; index--) {
                    cancellationToken.ThrowIfCancellationRequested();
                    pending.Push(array.Items[index]);
                }
                continue;
            }
            if (value is not PdfDictionary dictionary || !visitedContainers.Add(dictionary)) continue;
            foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                if (MatchName(item.Key)) return matchedGroups;
                pending.Push(item.Value);
            }
        }

        return matchedGroups;

        bool MatchName(string name) {
            for (int groupIndex = 0; groupIndex < requestedNameGroups.Count; groupIndex++) {
                int groupBit = 1 << groupIndex;
                if ((matchedGroups & groupBit) == 0 && requestedNameGroups[groupIndex].Contains(name)) {
                    matchedGroups |= groupBit;
                }
            }
            return matchedGroups == allGroups;
        }
    }

    private static bool ShouldSuppressParsedPdfNameException(Exception exception, PdfLoadOptions? options) {
        if (exception is OperationCanceledException || exception is OutOfMemoryException || exception is StackOverflowException) {
            return false;
        }

        return options is null || exception is not PdfEncryptionException;
    }

    private static bool VisitParsedPdfNames(PdfObject value, Func<string, bool> visit, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfName name:
                return visit(name.Name);
            case PdfDictionary dictionary:
                foreach (var item in dictionary.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (visit(item.Key) || VisitParsedPdfNames(item.Value, visit, cancellationToken)) {
                        return true;
                    }
                }

                return false;
            case PdfArray array:
                foreach (PdfObject item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (VisitParsedPdfNames(item, visit, cancellationToken)) {
                        return true;
                    }
                }

                return false;
            case PdfStream stream:
                return VisitParsedPdfNames(stream.Dictionary, visit, cancellationToken);
            default:
                return false;
        }
    }

}
