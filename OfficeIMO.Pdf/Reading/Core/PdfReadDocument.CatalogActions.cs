namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfCatalogAction> ExtractCatalogActions(out IReadOnlyList<PdfJavaScript> javaScripts, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            javaScripts = Array.Empty<PdfJavaScript>();
            return Array.Empty<PdfCatalogAction>();
        }

        PdfObject? javaScriptNameTree = null;
        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDict(namesObject) is PdfDictionary namesDictionary) {
            namesDictionary.Items.TryGetValue("JavaScript", out javaScriptNameTree);
        }
        if (javaScriptNameTree is null &&
            !catalog.Items.ContainsKey("OpenAction") &&
            !catalog.Items.ContainsKey("AA")) {
            javaScripts = Array.Empty<PdfJavaScript>();
            return Array.Empty<PdfCatalogAction>();
        }

        var result = new List<PdfCatalogAction>();
        var scripts = new List<PdfJavaScript>();
        if (javaScriptNameTree is not null) {
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
                ref totalJavaScriptBytes,
                cancellationToken);
        }

        if (catalog.Items.TryGetValue("OpenAction", out var openAction)) {
            AddCatalogAction("OpenAction", "OpenAction", null, openAction, result, new HashSet<int>(), cancellationToken);
        }

        if (catalog.Items.TryGetValue("AA", out var additionalActionsObject) &&
            ResolveObject(additionalActionsObject) is PdfDictionary additionalActions) {
            foreach (var item in additionalActions.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                AddCatalogAction("AA." + item.Key, "AA", item.Key, item.Value, result, new HashSet<int>(), cancellationToken);
            }
        }

        if (scripts.Count == 0) {
            javaScripts = Array.Empty<PdfJavaScript>();
        } else {
            var ordered = new List<(PdfJavaScript Script, int Index)>(scripts.Count);
            for (int index = 0; index < scripts.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                ordered.Add((scripts[index], index));
            }
            try {
                ordered.Sort((left, right) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    int comparison = StringComparer.Ordinal.Compare(left.Script.Name, right.Script.Name);
                    return comparison != 0 ? comparison : left.Index.CompareTo(right.Index);
                });
            } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
                cancellationToken.ThrowIfCancellationRequested();
                throw;
            }
            cancellationToken.ThrowIfCancellationRequested();
            var sorted = new List<PdfJavaScript>(ordered.Count);
            foreach (var entry in ordered) {
                cancellationToken.ThrowIfCancellationRequested();
                sorted.Add(entry.Script);
            }
            javaScripts = sorted.AsReadOnly();
        }
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
        ref long totalJavaScriptBytes,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();
                discoveredJavaScripts++;
                if (discoveredJavaScripts > _options.Limits.MaxJavaScripts) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, _options.Limits.MaxJavaScripts, discoveredJavaScripts);
                }
                if (TryReadCatalogActionName(actionNames.Items[i], out string? name, cancellationToken)) {
                    AddCatalogAction(name!, "Names/JavaScript", null, actionNames.Items[i + 1], result, new HashSet<int>(), cancellationToken);
                    bool hasReadableSource = TryReadJavaScriptSource(actionNames.Items[i + 1], out string? script, out long sourceBytes, cancellationToken);
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
                cancellationToken.ThrowIfCancellationRequested();
                AddCatalogActionsFromNameTree(
                    kid,
                    result,
                    scripts,
                    visitedReferences,
                    depth + 1,
                    ref traversedNodes,
                    ref discoveredJavaScripts,
                    ref totalJavaScriptBytes,
                    cancellationToken);
            }
        }
    }

    private bool TryReadJavaScriptSource(PdfObject actionObject, out string? script, out long sourceBytes, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
            bool decoded = PdfJavaScriptStringEncoding.TryDecode(text.RawBytes, out script!, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return decoded;
        }

        if (source is PdfStream stream) {
            try {
                byte[] decoded = _decodedStreamBudget.DecodeRequired(
                    stream,
                    _objects,
                    Math.Min(_options.Limits.MaxJavaScriptBytes, _options.Limits.MaxDecodedStreamBytes),
                    cancellationToken);
                sourceBytes = decoded.LongLength;
                bool readable = PdfJavaScriptStringEncoding.TryDecode(decoded, out script!, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                return readable;
            } catch (PdfReadLimitException) {
                throw;
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

    private bool TryReadCatalogActionName(PdfObject obj, out string? name, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (ResolveObject(obj)) {
            case PdfStringObj text:
                return PdfJavaScriptStringEncoding.TryDecode(text.RawBytes, out name!, cancellationToken) && !string.IsNullOrEmpty(name);
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
        System.Threading.CancellationToken cancellationToken,
        string? actionPath = null,
        bool isChainedAction = false) {
        cancellationToken.ThrowIfCancellationRequested();
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

        if (TryReadCatalogActionType(dictionary, out string? actionType)) {
            string? uri = string.Equals(actionType, "URI", StringComparison.Ordinal) ? TryReadText(dictionary, "URI") : null;
            result.Add(new PdfCatalogAction(name, actionType!, source, triggerName, actionPath ?? GetDefaultCatalogActionPath(name, source), isChainedAction, uri, PdfActionPayloadFingerprint.Create(dictionary, _objects, _options.Limits, cancellationToken)));
        }

        if (dictionary.Items.TryGetValue("Next", out var nextAction)) {
            AddCatalogNextActions(name + ".Next", source, triggerName, nextAction, result, pathReferences, cancellationToken);
        }
    }

    private void AddCatalogNextActions(
        string name,
        string source,
        string? triggerName,
        PdfObject obj,
        List<PdfCatalogAction> result,
        HashSet<int> visitedReferences,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();
                int before = result.Count;
                string nextPath = name + "." + activeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
                AddCatalogAction(nextPath, source, triggerName, actions.Items[i], result, new HashSet<int>(pathReferences), cancellationToken, nextPath, isChainedAction: true);
                if (result.Count > before) {
                    activeIndex++;
                }
            }

            return;
        }

        if (resolved is PdfDictionary) {
            AddCatalogAction(name, source, triggerName, resolved, result, pathReferences, cancellationToken, name, isChainedAction: true);
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

}
