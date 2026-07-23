namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private (int? PageNumber, double? DestinationTop, PdfOpenActionDestinationMode? DestinationMode, double? DestinationLeft, double? DestinationBottom, double? DestinationRight, double? DestinationZoom) GetOutlineDestination(PdfDictionary item) {
        if (item.Items.TryGetValue("Dest", out var destObj) &&
            TryReadDestinationOrNamedDestination(destObj, out int? pageNumber, out double? destinationTop, out PdfOpenActionDestinationMode? destinationMode, out double? destinationLeft, out double? destinationBottom, out double? destinationRight, out double? destinationZoom)) {
            return (pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight, destinationZoom);
        }

        if (item.Items.TryGetValue("A", out var actionObject) &&
            ResolveObject(actionObject) is PdfDictionary action &&
            action.Get<PdfName>("S")?.Name == "GoTo" &&
            action.Items.TryGetValue("D", out var actionDestination) &&
            TryReadDestinationOrNamedDestination(actionDestination, out pageNumber, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight, out destinationZoom)) {
            return (pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight, destinationZoom);
        }

        return (null, null, null, null, null, null, null);
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
            int traversedNameTreeNodes = 0;
            AddNamedDestinationsFromNameTree(namedDestinationTree, result, new HashSet<int>(), 0, ref traversedNameTreeNodes);
        }

        return result.Count == 0 ? Array.Empty<PdfNamedDestination>() : result.AsReadOnly();
    }

    private void AddNamedDestinationsFromNameTree(
        PdfObject treeObject,
        List<PdfNamedDestination> result,
        HashSet<int> visitedReferences,
        int depth,
        ref int traversedNodes) {
        EnsureNameTreeBudget(depth, traversedNodes);
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber)) {
                return;
            }

            EnsureNameTreeBudget(depth, ++traversedNodes);
            if (!PdfObjectLookup.TryGet(_objects, reference, out var indirect)) {
                return;
            }

            treeObject = indirect.Value;
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
                AddNamedDestinationsFromNameTree(kid, result, visitedReferences, depth + 1, ref traversedNodes);
            }
        }
    }

    private void EnsureNameTreeBudget(int depth, int traversedNodes) {
        if (depth > _options.Limits.MaxNameTreeDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeDepth, _options.Limits.MaxNameTreeDepth, depth);
        }

        if (traversedNodes > _options.Limits.MaxNameTreeNodes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeNodes, _options.Limits.MaxNameTreeNodes, traversedNodes);
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
        if (string.IsNullOrEmpty(name) ||
            !TryReadDestination(destinationObject, out int? pageNumber, out double? destinationTop, out PdfOpenActionDestinationMode? destinationMode, out double? destinationLeft, out double? destinationBottom, out double? destinationRight, out double? destinationZoom)) {
            return false;
        }

        destination = new PdfNamedDestination(name, pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight, destinationZoom);
        return true;
    }

    private bool TryReadDestinationOrNamedDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        return TryReadDestinationOrNamedDestination(destinationObject, out pageNumber, out destinationTop, out _);
    }

    private bool TryReadDestinationOrNamedDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop, out PdfOpenActionDestinationMode? destinationMode) {
        return TryReadDestinationOrNamedDestination(destinationObject, out pageNumber, out destinationTop, out destinationMode, out _, out _, out _);
    }

    private bool TryReadDestinationOrNamedDestination(
        PdfObject destinationObject,
        out int? pageNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight) {
        return TryReadDestinationOrNamedDestination(destinationObject, out pageNumber, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight, out _);
    }

    private bool TryReadDestinationOrNamedDestination(
        PdfObject destinationObject,
        out int? pageNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight,
        out double? destinationZoom) {
        if (TryReadDestination(destinationObject, out pageNumber, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight, out destinationZoom)) {
            return true;
        }

        if (TryReadDestinationName(destinationObject, out string? name, out var kind)) {
            var lookup = kind == PdfNamedDestinationTokenKind.String ? _stringDestinations : _nameDestinations;
            if (lookup.TryGetValue(name!, out var destination)) {
                pageNumber = destination.PageNumber;
                destinationTop = destination.DestinationTop;
                destinationMode = destination.DestinationMode;
                destinationLeft = destination.DestinationLeft;
                destinationBottom = destination.DestinationBottom;
                destinationRight = destination.DestinationRight;
                destinationZoom = destination.DestinationZoom;
                return true;
            }
        }

        pageNumber = null;
        destinationTop = null;
        destinationMode = null;
        destinationLeft = null;
        destinationBottom = null;
        destinationRight = null;
        destinationZoom = null;
        return false;
    }

    private bool TryReadDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop) {
        return TryReadDestination(destinationObject, out pageNumber, out destinationTop, out _);
    }

    private bool TryReadDestination(PdfObject destinationObject, out int? pageNumber, out double? destinationTop, out PdfOpenActionDestinationMode? destinationMode) {
        return TryReadDestination(destinationObject, out pageNumber, out destinationTop, out destinationMode, out _, out _, out _);
    }

    private bool TryReadDestination(
        PdfObject destinationObject,
        out int? pageNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight) {
        return TryReadDestination(destinationObject, out pageNumber, out destinationTop, out destinationMode, out destinationLeft, out destinationBottom, out destinationRight, out _);
    }

    private bool TryReadDestination(
        PdfObject destinationObject,
        out int? pageNumber,
        out double? destinationTop,
        out PdfOpenActionDestinationMode? destinationMode,
        out double? destinationLeft,
        out double? destinationBottom,
        out double? destinationRight,
        out double? destinationZoom) {
        pageNumber = null;
        destinationTop = null;
        destinationMode = null;
        destinationLeft = null;
        destinationBottom = null;
        destinationRight = null;
        destinationZoom = null;

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

        if (destination.Items.Count > 1 && ResolveObject(destination.Items[1]) is PdfName fitName) {
            switch (fitName.Name) {
                case "XYZ":
                    destinationMode = PdfOpenActionDestinationMode.Xyz;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber xyzLeft) {
                        destinationLeft = xyzLeft.Value;
                    }

                    if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber xyzTop) {
                        destinationTop = xyzTop.Value;
                    }

                    if (destination.Items.Count > 4 && ResolveObject(destination.Items[4]) is PdfNumber xyzZoom) {
                        destinationZoom = xyzZoom.Value;
                    }

                    break;
                case "Fit":
                    destinationMode = PdfOpenActionDestinationMode.Fit;
                    break;
                case "FitH":
                    destinationMode = PdfOpenActionDestinationMode.FitHorizontal;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitTop) {
                        destinationTop = fitTop.Value;
                    }

                    break;
                case "FitV":
                    destinationMode = PdfOpenActionDestinationMode.FitVertical;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitLeft) {
                        destinationLeft = fitLeft.Value;
                    }

                    break;
                case "FitR":
                    destinationMode = PdfOpenActionDestinationMode.FitRectangle;
                    if (destination.Items.Count > 5) {
                        if (ResolveObject(destination.Items[2]) is PdfNumber left) {
                            destinationLeft = left.Value;
                        }

                        if (ResolveObject(destination.Items[3]) is PdfNumber bottom) {
                            destinationBottom = bottom.Value;
                        }

                        if (ResolveObject(destination.Items[4]) is PdfNumber right) {
                            destinationRight = right.Value;
                        }

                        if (ResolveObject(destination.Items[5]) is PdfNumber top) {
                            destinationTop = top.Value;
                        }
                    }

                    break;
                case "FitB":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBox;
                    break;
                case "FitBH":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBoxHorizontal;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitBoundingTop) {
                        destinationTop = fitBoundingTop.Value;
                    }

                    break;
                case "FitBV":
                    destinationMode = PdfOpenActionDestinationMode.FitBoundingBoxVertical;
                    if (destination.Items.Count > 2 && ResolveObject(destination.Items[2]) is PdfNumber fitBoundingLeft) {
                        destinationLeft = fitBoundingLeft.Value;
                    }

                    break;
                default:
                    if (destination.Items.Count > 3 && ResolveObject(destination.Items[3]) is PdfNumber fallbackTop) {
                        destinationTop = fallbackTop.Value;
                    }

                    break;
            }
        }

        return true;
    }

    private enum PdfNamedDestinationTokenKind {
        None,
        Name,
        String
    }

    internal int? GetPageNumberForObject(int objectNumber) {
        for (int i = 0; i < Pages.Count; i++) {
            if (Pages[i].ObjectNumber == objectNumber) {
                return i + 1;
            }
        }

        return null;
    }
}
