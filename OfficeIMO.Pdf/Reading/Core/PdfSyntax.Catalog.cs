using System.Text.RegularExpressions;

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

    internal static PdfDictionary? FindCatalog(Dictionary<int, PdfIndirectObject> map, string? trailerRaw = null) {
        if (TryGetTrailerRootReference(trailerRaw, out PdfReference rootReference) &&
            PdfObjectLookup.TryGet(map, rootReference, out var rootObject) &&
            rootObject.Value is PdfDictionary rootDictionary &&
            rootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return rootDictionary;
        }

        if (TryGetXrefStreamRootReference(map, out rootReference) &&
            PdfObjectLookup.TryGet(map, rootReference, out rootObject) &&
            rootObject.Value is PdfDictionary xrefRootDictionary &&
            xrefRootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return xrefRootDictionary;
        }

        return FindCatalogByScan(map);
    }

    private static PdfDictionary? FindCatalogByScan(Dictionary<int, PdfIndirectObject> map) {
        foreach (var entry in map.Values) {
            if (entry.Value is PdfDictionary dictionary &&
                dictionary.Get<PdfName>("Type")?.Name == "Catalog") {
                return dictionary;
            }
        }

        return null;
    }

    private static bool TryGetTrailerRootReference(string? trailerRaw, out PdfReference reference) {
        reference = null!;
        if (string.IsNullOrWhiteSpace(trailerRaw)) {
            return false;
        }

        Match match = TrailerRootRegex.Match(trailerRaw);
        if (!match.Success ||
            !int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int objectNumber) ||
            !int.TryParse(match.Groups[2].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int generation)) {
            return false;
        }

        reference = new PdfReference(objectNumber, generation);
        return true;
    }

    private static bool TryGetXrefStreamRootReference(Dictionary<int, PdfIndirectObject> map, out PdfReference reference) {
        reference = null!;
        foreach (var entry in map.Values.OrderByDescending(static item => item.ObjectNumber)) {
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
                return true;
            }
        }

        return false;
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

    private static bool ContainsAnyParsedPdfName(byte[] pdf, params string[] names) {
        return ContainsAnyParsedPdfName(pdf, null, names);
    }

    private static bool ContainsAnyParsedPdfName(byte[] pdf, PdfReadOptions? options, params string[] names) {
        try {
            var (map, _) = ParseObjects(pdf, options);
            var nameSet = new HashSet<string>(names, StringComparer.Ordinal);
            foreach (PdfIndirectObject indirectObject in map.Values) {
                if (ContainsAnyParsedPdfName(indirectObject.Value, nameSet)) {
                    return true;
                }
            }
        } catch (Exception ex) when (ShouldSuppressParsedPdfNameException(ex, options)) {
            return false;
        }

        return false;
    }

    internal static bool ContainsAnyParsedPdfName(
        IReadOnlyDictionary<int, PdfIndirectObject> objects,
        params string[] names) {
        var nameSet = new HashSet<string>(names, StringComparer.Ordinal);
        foreach (PdfIndirectObject indirectObject in objects.Values) {
            if (ContainsAnyParsedPdfName(indirectObject.Value, nameSet)) {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldSuppressParsedPdfNameException(Exception exception, PdfReadOptions? options) {
        if (exception is OutOfMemoryException || exception is StackOverflowException) {
            return false;
        }

        return options is null || exception is not PdfEncryptionException;
    }

    private static bool ContainsAnyParsedPdfName(PdfObject value, HashSet<string> names) {
        switch (value) {
            case PdfName name:
                return names.Contains(name.Name);
            case PdfDictionary dictionary:
                foreach (var item in dictionary.Items) {
                    if (names.Contains(item.Key) || ContainsAnyParsedPdfName(item.Value, names)) {
                        return true;
                    }
                }

                return false;
            case PdfArray array:
                foreach (PdfObject item in array.Items) {
                    if (ContainsAnyParsedPdfName(item, names)) {
                        return true;
                    }
                }

                return false;
            case PdfStream stream:
                return ContainsAnyParsedPdfName(stream.Dictionary, names);
            default:
                return false;
        }
    }

}
