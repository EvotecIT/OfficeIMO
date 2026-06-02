namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private (int? PageNumber, double? DestinationTop) GetOutlineDestination(PdfDictionary item) {
        if (item.Items.TryGetValue("Dest", out var destObj) &&
            TryReadDestinationOrNamedDestination(destObj, out int? pageNumber, out double? destinationTop)) {
            return (pageNumber, destinationTop);
        }

        if (item.Items.TryGetValue("A", out var actionObject) &&
            ResolveObject(actionObject) is PdfDictionary action &&
            action.Get<PdfName>("S")?.Name == "GoTo" &&
            action.Items.TryGetValue("D", out var actionDestination) &&
            TryReadDestinationOrNamedDestination(actionDestination, out pageNumber, out destinationTop)) {
            return (pageNumber, destinationTop);
        }

        return (null, null);
    }

    private IReadOnlyList<PdfNamedDestination> ExtractNamedDestinations() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            return Array.Empty<PdfNamedDestination>();
        }

        var result = new List<PdfNamedDestination>();
        if (catalog.Items.TryGetValue("Dests", out var directDests) &&
            ResolveDict(directDests) is PdfDictionary directDestinations) {
            foreach (var entry in directDestinations.Items) {
                if (TryCreateNamedDestination(entry.Key, entry.Value, out var destination)) {
                    AddNamedDestination(result, destination, PdfNamedDestinationTokenKind.Name);
                }
            }
        }

        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDict(namesObject) is PdfDictionary namesDictionary &&
            namesDictionary.Items.TryGetValue("Dests", out var namedDestinationTree)) {
            AddNamedDestinationsFromNameTree(namedDestinationTree, result, new HashSet<int>());
        }

        return result.Count == 0 ? Array.Empty<PdfNamedDestination>() : result.AsReadOnly();
    }

    private void AddNamedDestinationsFromNameTree(
        PdfObject treeObject,
        List<PdfNamedDestination> result,
        HashSet<int> visitedReferences) {
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(_objects, reference, out var indirect)) {
                return;
            }

            AddNamedDestinationsFromNameTree(indirect.Value, result, visitedReferences);
            return;
        }

        if (treeObject is not PdfDictionary tree) {
            return;
        }

        if (tree.Items.TryGetValue("Names", out var destinationNamesObject) &&
            ResolveArray(destinationNamesObject) is PdfArray destinationNames) {
            for (int i = 0; i + 1 < destinationNames.Items.Count; i += 2) {
                if (TryReadDestinationName(destinationNames.Items[i], out string? name, out _) &&
                    TryCreateNamedDestination(name!, destinationNames.Items[i + 1], out var destination)) {
                    AddNamedDestination(result, destination, PdfNamedDestinationTokenKind.String);
                }
            }
        }

        if (tree.Items.TryGetValue("Kids", out var kidsObject) &&
            ResolveArray(kidsObject) is PdfArray kids) {
            foreach (var kid in kids.Items) {
                AddNamedDestinationsFromNameTree(kid, result, visitedReferences);
            }
        }
    }

    private void AddNamedDestination(List<PdfNamedDestination> result, PdfNamedDestination destination, PdfNamedDestinationTokenKind kind) {
        result.Add(destination);
        var lookup = kind == PdfNamedDestinationTokenKind.String ? _stringDestinations : _nameDestinations;
#if NETSTANDARD2_0 || NET472
        if (!lookup.ContainsKey(destination.Name)) {
            lookup[destination.Name] = destination;
        }
#else
        lookup.TryAdd(destination.Name, destination);
#endif
    }

    private bool TryReadDestinationName(PdfObject obj, out string? name, out PdfNamedDestinationTokenKind kind) {
        PdfObject? resolved = ResolveObject(obj);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        switch (resolved) {
            case PdfStringObj text:
                name = text.Value;
                kind = PdfNamedDestinationTokenKind.String;
                return !string.IsNullOrEmpty(name);
            case PdfName pdfName:
                name = pdfName.Name;
                kind = PdfNamedDestinationTokenKind.Name;
                return !string.IsNullOrEmpty(name);
            default:
                name = null;
                kind = PdfNamedDestinationTokenKind.None;
                return false;
        }
    }

    private bool TryCreateNamedDestination(string name, PdfObject destinationObject, out PdfNamedDestination destination) {
        destination = null!;
        if (string.IsNullOrEmpty(name) || !TryReadDestination(destinationObject, out int? pageNumber, out double? destinationTop)) {
            return false;
        }

        destination = new PdfNamedDestination(name, pageNumber, destinationTop);
        return true;
    }

    private bool TryReadDestinationOrNamedDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        if (TryReadDestination(destinationObject, out pageNumber, out destinationTop)) {
            return true;
        }

        if (TryReadDestinationName(destinationObject, out string? name, out var kind)) {
            var lookup = kind == PdfNamedDestinationTokenKind.String ? _stringDestinations : _nameDestinations;
            if (lookup.TryGetValue(name!, out var destination)) {
                pageNumber = destination.PageNumber;
                destinationTop = destination.DestinationTop;
                return true;
            }
        }

        pageNumber = null;
        destinationTop = null;
        return false;
    }

    private bool TryReadDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        pageNumber = null;
        destinationTop = null;

        PdfObject? resolved = ResolveObject(destinationObject);
        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            resolved = ResolveObject(explicitDestination);
        }

        if (resolved is not PdfArray destination || destination.Items.Count == 0) {
            return false;
        }

        if (destination.Items[0] is PdfReference pageRef) {
            pageNumber = GetPageNumberForObject(pageRef.ObjectNumber);
        }

        if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber top) {
            destinationTop = top.Value;
        }

        return true;
    }

    private enum PdfNamedDestinationTokenKind {
        None,
        Name,
        String
    }

    private int? GetPageNumberForObject(int objectNumber) {
        for (int i = 0; i < Pages.Count; i++) {
            if (Pages[i].ObjectNumber == objectNumber) {
                return i + 1;
            }
        }

        return null;
    }
}
