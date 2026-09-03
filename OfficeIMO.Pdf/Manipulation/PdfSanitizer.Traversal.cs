namespace OfficeIMO.Pdf;

internal static partial class PdfSanitizer {
    private static readonly HashSet<string> RichAnnotationSubtypes = new HashSet<string>(StringComparer.Ordinal) {
        "RichMedia", "Movie", "Sound", "Screen", "3D", "FileAttachment"
    };

    private static IReadOnlyList<PdfSanitizationFinding> Scan(
        Dictionary<int, PdfIndirectObject> objects,
        PdfSanitizationOptions policy) {
        var findings = new List<PdfSanitizationFinding>();
        foreach (KeyValuePair<int, PdfIndirectObject> item in objects.OrderBy(static item => item.Key)) {
            policy.CancellationToken.ThrowIfCancellationRequested();
            ScanObject(objects, item.Value.Value, policy, item.Key, "Object[" + item.Key + "]", findings);
        }

        return findings.Count == 0 ? Array.Empty<PdfSanitizationFinding>() : findings.AsReadOnly();
    }

    private static void ScanObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfSanitizationOptions policy,
        int objectNumber,
        string path,
        List<PdfSanitizationFinding> findings) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        if (value is PdfStream stream) {
            ScanDictionary(objects, stream.Dictionary, policy, objectNumber, path, findings);
        } else if (value is PdfDictionary dictionary) {
            ScanDictionary(objects, dictionary, policy, objectNumber, path, findings);
        } else if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                if (array.Items[i] is not PdfReference) {
                    ScanObject(objects, array.Items[i], policy, objectNumber, path + "[" + i + "]", findings);
                }
            }
        }
    }

    private static void ScanDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        PdfSanitizationOptions policy,
        int objectNumber,
        string path,
        List<PdfSanitizationFinding> findings) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        if (TryGetForbiddenAction(objects, dictionary, policy, out PdfSanitizationFindingKind findingKind, out PdfSanitizationActionKind actionKind, out string? actionDetail)) {
            findings.Add(new PdfSanitizationFinding(findingKind, objectNumber, path, actionDetail!, actionKind));
        }

        if (IsRichAnnotation(objects, dictionary, policy, out string? annotationSubtype)) {
            findings.Add(new PdfSanitizationFinding(PdfSanitizationFindingKind.RichMedia, objectNumber, path, annotationSubtype!));
        }

        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
            policy.CancellationToken.ThrowIfCancellationRequested();
            string itemPath = path + "/" + item.Key;
            if (item.Key == "EmbeddedFiles" || item.Key == "AF" || item.Key == "EF") {
                findings.Add(new PdfSanitizationFinding(PdfSanitizationFindingKind.EmbeddedFile, objectNumber, itemPath, item.Key));
            }

            if (item.Key == "URI" && Resolve(objects, item.Value) is PdfDictionary uriDictionary &&
                TryGetString(objects, uriDictionary, "Base", out string? baseUri) && policy.ShouldRemoveCatalogUriBase(baseUri!)) {
                PdfSanitizationFindingKind uriFindingKind = policy.ActionKindsToRemove.HasValue
                    ? PdfSanitizationFindingKind.ActiveAction
                    : PdfSanitizationFindingKind.UnsafeUri;
                findings.Add(new PdfSanitizationFinding(uriFindingKind, objectNumber, itemPath + "/Base", baseUri!, PdfSanitizationActionKind.Uri));
            }

            if (item.Value is not PdfReference) {
                ScanObject(objects, item.Value, policy, objectNumber, itemPath, findings);
            }
        }
    }

    private static void SanitizeObjectGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        int maximumActionNodes) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        var actionBudget = new PdfSanitizerActionBudget(maximumActionNodes);
        foreach (PdfIndirectObject item in objects.Values.OrderBy(static item => item.ObjectNumber)) {
            policy.CancellationToken.ThrowIfCancellationRequested();
            SanitizeObject(objects, item.Value, policy, maximumActionDepth, actionBudget);
        }

        foreach (PdfIndirectObject item in objects.Values.OrderBy(static item => item.ObjectNumber)) {
            policy.CancellationToken.ThrowIfCancellationRequested();
            RemoveEmptyContainers(objects, item.Value);
        }
    }

    private static void SanitizeObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        PdfSanitizerActionBudget actionBudget) {
        if (value is PdfStream stream) {
            SanitizeDictionary(objects, stream.Dictionary, policy, maximumActionDepth, actionBudget);
        } else if (value is PdfDictionary dictionary) {
            SanitizeDictionary(objects, dictionary, policy, maximumActionDepth, actionBudget);
        } else if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                if (array.Items[i] is not PdfReference) {
                    SanitizeObject(objects, array.Items[i], policy, maximumActionDepth, actionBudget);
                }
            }
        }
    }

    private static void SanitizeDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        PdfSanitizerActionBudget actionBudget) {
        if (policy.ShouldRemoveAction("JavaScript")) {
            dictionary.Items.Remove("JavaScript");
        }

        dictionary.Items.Remove("EmbeddedFiles");
        dictionary.Items.Remove("AF");
        dictionary.Items.Remove("EF");
        bool actionTraversalAlreadyNormalized = actionBudget.WasNormalized(dictionary);

        if (dictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) &&
            Resolve(objects, annotationsObject) is PdfArray annotations) {
            FilterAnnotations(objects, annotations, policy);
        }

        string[] keys = dictionary.Items.Keys.ToArray();
        for (int i = 0; i < keys.Length; i++) {
            string key = keys[i];
            if (!dictionary.Items.TryGetValue(key, out PdfObject? item)) {
                continue;
            }
            if (actionTraversalAlreadyNormalized && string.Equals(key, "Next", StringComparison.Ordinal)) continue;

            PdfObject? resolved = Resolve(objects, item);
            if (resolved is PdfDictionary action && TryGetForbiddenAction(objects, action, policy, out _, out _, out _)) {
                List<PdfDictionary> retained = action.Items.TryGetValue("Next", out PdfObject? next)
                    ? CollectRetainedActions(objects, next, policy, maximumActionDepth, 0, new HashSet<(int ObjectNumber, int Generation)>(), actionBudget)
                    : new List<PdfDictionary>();
                if (retained.Count == 0) {
                    dictionary.Items.Remove(key);
                } else {
                    PdfDictionary promoted = CreatePromotedActionRoot(retained);
                    SanitizeNormalizedActionDictionary(objects, promoted, policy, maximumActionDepth, actionBudget);
                    dictionary.Items[key] = promoted;
                }
                continue;
            }

            if (key == "Next" && resolved is PdfArray nextActions) {
                FilterActions(objects, nextActions, policy, maximumActionDepth, actionBudget);
                SanitizeNormalizedActionArray(objects, nextActions, policy, maximumActionDepth, actionBudget);
                continue;
            }

            if (key == "URI" && resolved is PdfDictionary uriDictionary &&
                TryGetString(objects, uriDictionary, "Base", out string? baseUri) && policy.ShouldRemoveCatalogUriBase(baseUri!)) {
                uriDictionary.Items.Remove("Base");
            }

            if (item is not PdfReference) {
                SanitizeObject(objects, item, policy, maximumActionDepth, actionBudget);
            }
        }
    }

    private static void SanitizeNormalizedActionArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray actions,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        PdfSanitizerActionBudget actionBudget) {
        for (int i = 0; i < actions.Items.Count; i++) {
            if (actions.Items[i] is PdfDictionary action) {
                SanitizeNormalizedActionDictionary(objects, action, policy, maximumActionDepth, actionBudget);
            }
        }
    }

    private static void SanitizeNormalizedActionDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary action,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        PdfSanitizerActionBudget actionBudget) {
        if (action.Items.TryGetValue("Next", out PdfObject? nextObject)) {
            if (nextObject is PdfDictionary nextAction) {
                SanitizeNormalizedActionDictionary(objects, nextAction, policy, maximumActionDepth, actionBudget);
            } else if (nextObject is PdfArray nextActions) {
                SanitizeNormalizedActionArray(objects, nextActions, policy, maximumActionDepth, actionBudget);
            }
        }

        foreach (KeyValuePair<string, PdfObject> item in action.Items.ToArray()) {
            if (string.Equals(item.Key, "Next", StringComparison.Ordinal) || item.Value is PdfReference) continue;
            SanitizeObject(objects, item.Value, policy, maximumActionDepth, actionBudget);
        }
    }

    private static List<PdfDictionary> CollectRetainedActions(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject actionObject,
        PdfSanitizationOptions policy,
        int maximumDepth,
        int depth,
        HashSet<(int ObjectNumber, int Generation)> pathReferences,
        PdfSanitizerActionBudget actionBudget) {
        if (depth > maximumDepth) throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, maximumDepth, depth);

        if (actionObject is PdfReference reference) {
            var key = (reference.ObjectNumber, reference.Generation);
            if (!pathReferences.Add(key) || !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return new List<PdfDictionary>();
            var nextPath = new HashSet<(int ObjectNumber, int Generation)>(pathReferences);
            return CollectRetainedActions(objects, indirect.Value, policy, maximumDepth, depth + 1, nextPath, actionBudget);
        }

        if (actionObject is PdfArray array) {
            var retained = new List<PdfDictionary>();
            for (int i = 0; i < array.Items.Count; i++) {
                retained.AddRange(CollectRetainedActions(objects, array.Items[i], policy, maximumDepth, depth + 1, new HashSet<(int ObjectNumber, int Generation)>(pathReferences), actionBudget));
            }
            return retained;
        }

        if (actionObject is not PdfDictionary action) return new List<PdfDictionary>();
        actionBudget.Consume();
        List<PdfDictionary> children = action.Items.TryGetValue("Next", out PdfObject? nextObject)
            ? CollectRetainedActions(objects, nextObject, policy, maximumDepth, depth + 1, new HashSet<(int ObjectNumber, int Generation)>(pathReferences), actionBudget)
            : new List<PdfDictionary>();
        action.Items.Remove("Next");
        AttachNextActions(action, children);
        actionBudget.MarkNormalized(action);
        if (TryGetForbiddenAction(objects, action, policy, out _, out _, out _)) return children;

        var clone = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> item in action.Items) {
            if (!string.Equals(item.Key, "Next", StringComparison.Ordinal)) clone.Items[item.Key] = item.Value;
        }
        AttachNextActions(clone, children);
        return new List<PdfDictionary> { clone };
    }

    private static PdfDictionary CreatePromotedActionRoot(List<PdfDictionary> retained) {
        PdfDictionary root = retained[0];
        if (retained.Count == 1) return root;
        var siblings = new List<PdfDictionary>(retained.Count - 1);
        for (int i = 1; i < retained.Count; i++) siblings.Add(retained[i]);
        AppendNextActions(root, siblings);
        return root;
    }

    private static void AttachNextActions(PdfDictionary action, List<PdfDictionary> nextActions) {
        if (nextActions.Count == 0) return;
        if (nextActions.Count == 1) {
            action.Items["Next"] = nextActions[0];
            return;
        }
        var array = new PdfArray();
        for (int i = 0; i < nextActions.Count; i++) array.Items.Add(nextActions[i]);
        action.Items["Next"] = array;
    }

    private static void AppendNextActions(PdfDictionary action, List<PdfDictionary> additionalActions) {
        if (additionalActions.Count == 0) return;
        if (!action.Items.TryGetValue("Next", out PdfObject? existing)) {
            AttachNextActions(action, additionalActions);
            return;
        }
        var combined = new PdfArray();
        if (existing is PdfArray existingArray) {
            for (int i = 0; i < existingArray.Items.Count; i++) combined.Items.Add(existingArray.Items[i]);
        } else {
            combined.Items.Add(existing);
        }
        for (int i = 0; i < additionalActions.Count; i++) combined.Items.Add(additionalActions[i]);
        action.Items["Next"] = combined;
    }

    private static void FilterActions(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray actions,
        PdfSanitizationOptions policy,
        int maximumActionDepth,
        PdfSanitizerActionBudget actionBudget) {
        var retained = new List<PdfDictionary>();
        for (int i = 0; i < actions.Items.Count; i++) {
            retained.AddRange(CollectRetainedActions(
                objects,
                actions.Items[i],
                policy,
                maximumActionDepth,
                depth: 0,
                new HashSet<(int ObjectNumber, int Generation)>(),
                actionBudget));
        }
        actions.Items.Clear();
        for (int i = 0; i < retained.Count; i++) actions.Items.Add(retained[i]);
    }

    private sealed class PdfSanitizerActionBudget {
        private readonly int _limit;
        private readonly HashSet<PdfDictionary> _normalizedActions = new();
        private int _count;

        internal PdfSanitizerActionBudget(int limit) {
            _limit = limit;
        }

        internal void Consume() {
            _count++;
            if (_count > _limit) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.IndirectObjects, _limit, _count);
            }
        }

        internal void MarkNormalized(PdfDictionary action) => _normalizedActions.Add(action);

        internal bool WasNormalized(PdfDictionary action) => _normalizedActions.Contains(action);
    }

    private static void FilterAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray annotations,
        PdfSanitizationOptions policy) {
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            if (Resolve(objects, annotations.Items[i]) is PdfDictionary annotation &&
                IsRichAnnotation(objects, annotation, policy, out _)) {
                annotations.Items.RemoveAt(i);
            }
        }
    }

    private static void RemoveEmptyContainers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value) {
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary is null) {
            return;
        }

        RemoveEmptyDictionary(objects, dictionary, "AA");
        RemoveEmptyDictionary(objects, dictionary, "Names");
        RemoveEmptyArray(objects, dictionary, "Annots");
        RemoveEmptyArray(objects, dictionary, "Next");
    }

    private static void RemoveEmptyDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string key) {
        if (owner.Items.TryGetValue(key, out PdfObject? value) &&
            Resolve(objects, value) is PdfDictionary dictionary &&
            dictionary.Items.Count == 0) {
            owner.Items.Remove(key);
        }
    }

    private static void RemoveEmptyArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string key) {
        if (owner.Items.TryGetValue(key, out PdfObject? value) &&
            Resolve(objects, value) is PdfArray array &&
            array.Items.Count == 0) {
            owner.Items.Remove(key);
        }
    }

    private static bool TryGetForbiddenAction(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        PdfSanitizationOptions policy,
        out PdfSanitizationFindingKind kind,
        out PdfSanitizationActionKind actionKind,
        out string? detail) {
        kind = PdfSanitizationFindingKind.ActiveAction;
        actionKind = PdfSanitizationActionKind.None;
        detail = null;
        if (Resolve(objects, dictionary.Get<PdfObject>("S")) is not PdfName actionName) {
            return false;
        }

        string actionType = actionName.Name;
        actionKind = PdfSanitizationOptions.GetActionKind(actionType);
        if (actionType == "URI") {
            bool hasUri = TryGetString(objects, dictionary, "URI", out string? uri);
            if (policy.ShouldRemoveAction(actionType, hasUri ? uri : null)) {
                kind = policy.ActionKindsToRemove.HasValue
                    ? PdfSanitizationFindingKind.ActiveAction
                    : PdfSanitizationFindingKind.UnsafeUri;
                detail = hasUri ? uri : actionType;
                return true;
            }

            return false;
        }

        if (!policy.ShouldRemoveAction(actionType)) {
            return false;
        }

        detail = actionType;
        return true;
    }

    private static bool IsRichAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        PdfSanitizationOptions policy,
        out string? subtype) {
        subtype = null;
        if (!policy.RemoveRichMedia || Resolve(objects, dictionary.Get<PdfObject>("Subtype")) is not PdfName name) {
            return false;
        }

        subtype = name.Name;
        return RichAnnotationSubtypes.Contains(subtype);
    }

    private static bool TryGetString(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        out string? value) {
        if (dictionary.Items.TryGetValue(key, out PdfObject? item) && Resolve(objects, item) is PdfStringObj text) {
            value = text.Value;
            return true;
        }

        value = null;
        return false;
    }

    private static PdfObject? Resolve(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return PdfObjectLookup.Resolve(objects, value);
    }
}
