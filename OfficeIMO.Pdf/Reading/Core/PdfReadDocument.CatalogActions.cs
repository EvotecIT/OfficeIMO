namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfCatalogAction> ExtractCatalogActions(out IReadOnlyList<PdfJavaScript> javaScripts) {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            javaScripts = Array.Empty<PdfJavaScript>();
            return Array.Empty<PdfCatalogAction>();
        }

        var result = new List<PdfCatalogAction>();
        var scripts = new List<PdfJavaScript>();
        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDict(namesObject) is PdfDictionary namesDictionary &&
            namesDictionary.Items.TryGetValue("JavaScript", out var javaScriptNameTree)) {
            int traversedNameTreeNodes = 0;
            int discoveredJavaScripts = 0;
            long totalJavaScriptBytes = 0L;
            AddCatalogActionsFromNameTree(
                javaScriptNameTree,
                result,
                scripts,
                new HashSet<(int ObjectNumber, int Generation)>(),
                0,
                ref traversedNameTreeNodes,
                ref discoveredJavaScripts,
                ref totalJavaScriptBytes);
        }

        if (catalog.Items.TryGetValue("OpenAction", out var openAction)) {
            AddCatalogAction("OpenAction", "OpenAction", null, openAction, result, new HashSet<int>());
        }

        if (catalog.Items.TryGetValue("AA", out var additionalActionsObject) &&
            ResolveObject(additionalActionsObject) is PdfDictionary additionalActions) {
            foreach (var item in additionalActions.Items) {
                AddCatalogAction("AA." + item.Key, "AA", item.Key, item.Value, result, new HashSet<int>());
            }
        }

        javaScripts = scripts.Count == 0
            ? Array.Empty<PdfJavaScript>()
            : scripts.OrderBy(static script => script.Name, StringComparer.Ordinal).ToList().AsReadOnly();
        return result.Count == 0 ? Array.Empty<PdfCatalogAction>() : result.AsReadOnly();
    }

    private void AddCatalogActionsFromNameTree(
        PdfObject treeObject,
        List<PdfCatalogAction> result,
        List<PdfJavaScript> scripts,
        HashSet<(int ObjectNumber, int Generation)> visitedReferences,
        int depth,
        ref int traversedNodes,
        ref int discoveredJavaScripts,
        ref long totalJavaScriptBytes) {
        EnsureNameTreeBudget(depth, traversedNodes);
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation))) {
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

        if (tree.Items.TryGetValue("Names", out var actionNamesObject) &&
            ResolveArray(actionNamesObject) is PdfArray actionNames) {
            for (int i = 0; i + 1 < actionNames.Items.Count; i += 2) {
                discoveredJavaScripts++;
                if (discoveredJavaScripts > _options.Limits.MaxJavaScripts) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, _options.Limits.MaxJavaScripts, discoveredJavaScripts);
                }
                if (TryReadCatalogActionName(actionNames.Items[i], out string? name)) {
                    AddCatalogAction(name!, "Names/JavaScript", null, actionNames.Items[i + 1], result, new HashSet<int>());
                    bool hasReadableSource = TryReadJavaScriptSource(actionNames.Items[i + 1], out string? script, out long sourceBytes);
                    totalJavaScriptBytes = checked(totalJavaScriptBytes + sourceBytes);
                    if (totalJavaScriptBytes > _options.Limits.MaxTotalJavaScriptBytes) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScriptBytes, _options.Limits.MaxTotalJavaScriptBytes, totalJavaScriptBytes);
                    }
                    if (hasReadableSource) {
                        scripts.Add(new PdfJavaScript(name!, script!));
                    }
                }
            }
        }

        if (tree.Items.TryGetValue("Kids", out var kidsObject) &&
            ResolveArray(kidsObject) is PdfArray kids) {
            foreach (var kid in kids.Items) {
                AddCatalogActionsFromNameTree(
                    kid,
                    result,
                    scripts,
                    visitedReferences,
                    depth + 1,
                    ref traversedNodes,
                    ref discoveredJavaScripts,
                    ref totalJavaScriptBytes);
            }
        }
    }

    private bool TryReadJavaScriptSource(PdfObject actionObject, out string? script, out long sourceBytes) {
        if (ResolveObject(actionObject) is not PdfDictionary action ||
            !TryReadCatalogActionType(action, out string? actionType) ||
            !string.Equals(actionType, "JavaScript", StringComparison.Ordinal) ||
            !action.Items.TryGetValue("JS", out PdfObject? sourceObject)) {
            script = null;
            sourceBytes = 0L;
            return false;
        }

        PdfObject? source = ResolveObject(sourceObject);
        if (source is PdfStringObj text) {
            int byteCount = text.RawBytes.Length;
            int maximumBytes = Math.Min(_options.Limits.MaxJavaScriptBytes, _options.Limits.MaxDecodedStreamBytes);
            if (byteCount > maximumBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumBytes, byteCount);
            }
            sourceBytes = byteCount;
            return PdfJavaScriptStringEncoding.TryDecode(text.RawBytes, out script!);
        }

        if (source is PdfStream stream) {
            try {
                byte[] decoded = Filters.StreamDecoder.DecodeRequired(
                    stream.Dictionary,
                    stream.Data,
                    _objects,
                    Math.Min(_options.Limits.MaxJavaScriptBytes, _options.Limits.MaxDecodedStreamBytes));
                sourceBytes = decoded.LongLength;
                return PdfJavaScriptStringEncoding.TryDecode(decoded, out script!);
            } catch (InvalidDataException) {
                script = null;
                sourceBytes = 0L;
                return false;
            }
        }

        script = null;
        sourceBytes = 0L;
        return false;
    }

    private bool TryReadCatalogActionName(PdfObject obj, out string? name) {
        switch (ResolveObject(obj)) {
            case PdfStringObj text:
                return PdfJavaScriptStringEncoding.TryDecode(text.RawBytes, out name!) && !string.IsNullOrEmpty(name);
            case PdfName pdfName:
                name = pdfName.Name;
                return !string.IsNullOrEmpty(name);
            default:
                name = null;
                return false;
        }
    }

    private void AddCatalogAction(
        string name,
        string source,
        string? triggerName,
        PdfObject obj,
        List<PdfCatalogAction> result,
        HashSet<int> visitedReferences,
        string? actionPath = null,
        bool isChainedAction = false) {
        HashSet<int> pathReferences = visitedReferences;
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference) {
            if (visitedReferences.Contains(reference.ObjectNumber)) {
                return;
            }

            pathReferences = new HashSet<int>(visitedReferences) { reference.ObjectNumber };
        }

        if (resolved is not PdfDictionary dictionary) {
            return;
        }

        if (TryReadCatalogActionType(dictionary, out string? actionType) &&
            IsActiveCatalogActionType(actionType!)) {
            string? uri = string.Equals(actionType, "URI", StringComparison.Ordinal) ? TryReadText(dictionary, "URI") : null;
            result.Add(new PdfCatalogAction(name, actionType!, source, triggerName, actionPath ?? GetDefaultCatalogActionPath(name, source), isChainedAction, uri));
        }

        if (dictionary.Items.TryGetValue("Next", out var nextAction)) {
            AddCatalogNextActions(name + ".Next", source, triggerName, nextAction, result, pathReferences);
        }
    }

    private void AddCatalogNextActions(
        string name,
        string source,
        string? triggerName,
        PdfObject obj,
        List<PdfCatalogAction> result,
        HashSet<int> visitedReferences) {
        HashSet<int> pathReferences = visitedReferences;
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference) {
            if (visitedReferences.Contains(reference.ObjectNumber)) {
                return;
            }

            pathReferences = new HashSet<int>(visitedReferences) { reference.ObjectNumber };
        }

        if (resolved is PdfArray actions) {
            int activeIndex = 0;
            for (int i = 0; i < actions.Items.Count; i++) {
                int before = result.Count;
                string nextPath = name + "." + activeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddCatalogAction(nextPath, source, triggerName, actions.Items[i], result, new HashSet<int>(pathReferences), nextPath, isChainedAction: true);
                if (result.Count > before) {
                    activeIndex++;
                }
            }

            return;
        }

        if (resolved is PdfDictionary) {
            AddCatalogAction(name, source, triggerName, resolved, result, pathReferences, name, isChainedAction: true);
        }
    }

    private static string? GetDefaultCatalogActionPath(string name, string source) {
        if (string.Equals(source, "AA", StringComparison.Ordinal) ||
            string.Equals(source, "OpenAction", StringComparison.Ordinal)) {
            return name;
        }

        return null;
    }

    private bool TryReadCatalogActionType(PdfDictionary dictionary, out string? actionType) {
        if (dictionary.Items.TryGetValue("S", out var actionTypeObject) &&
            ResolveObject(actionTypeObject) is PdfName pdfName &&
            !string.IsNullOrEmpty(pdfName.Name)) {
            actionType = pdfName.Name;
            return true;
        }

        actionType = null;
        return false;
    }

    private static bool IsActiveCatalogActionType(string actionType) {
        switch (actionType) {
            case "JavaScript":
            case "Launch":
            case "SubmitForm":
            case "ImportData":
            case "Movie":
            case "Rendition":
            case "RichMedia":
            case "URI":
                return true;
            default:
                return false;
        }
    }
}
