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
        if (TryGetForbiddenAction(objects, dictionary, policy, out PdfSanitizationFindingKind actionKind, out string? actionDetail)) {
            findings.Add(new PdfSanitizationFinding(actionKind, objectNumber, path, actionDetail!));
        }

        if (IsRichAnnotation(objects, dictionary, policy, out string? annotationSubtype)) {
            findings.Add(new PdfSanitizationFinding(PdfSanitizationFindingKind.RichMedia, objectNumber, path, annotationSubtype!));
        }

        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
            string itemPath = path + "/" + item.Key;
            if (item.Key == "EmbeddedFiles" || item.Key == "AF" || item.Key == "EF") {
                findings.Add(new PdfSanitizationFinding(PdfSanitizationFindingKind.EmbeddedFile, objectNumber, itemPath, item.Key));
            }

            if (item.Key == "URI" && Resolve(objects, item.Value) is PdfDictionary uriDictionary &&
                TryGetString(objects, uriDictionary, "Base", out string? baseUri) && !policy.IsUriAllowed(baseUri!)) {
                findings.Add(new PdfSanitizationFinding(PdfSanitizationFindingKind.UnsafeUri, objectNumber, itemPath + "/Base", baseUri!));
            }

            if (item.Value is not PdfReference) {
                ScanObject(objects, item.Value, policy, objectNumber, itemPath, findings);
            }
        }
    }

    private static void SanitizeObjectGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfSanitizationOptions policy) {
        foreach (PdfIndirectObject item in objects.Values.OrderBy(static item => item.ObjectNumber)) {
            SanitizeObject(objects, item.Value, policy);
        }

        foreach (PdfIndirectObject item in objects.Values.OrderBy(static item => item.ObjectNumber)) {
            RemoveEmptyContainers(objects, item.Value);
        }
    }

    private static void SanitizeObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        PdfSanitizationOptions policy) {
        if (value is PdfStream stream) {
            SanitizeDictionary(objects, stream.Dictionary, policy);
        } else if (value is PdfDictionary dictionary) {
            SanitizeDictionary(objects, dictionary, policy);
        } else if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                if (array.Items[i] is not PdfReference) {
                    SanitizeObject(objects, array.Items[i], policy);
                }
            }
        }
    }

    private static void SanitizeDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        PdfSanitizationOptions policy) {
        if (!policy.IsActionAllowed("JavaScript")) {
            dictionary.Items.Remove("JavaScript");
        }

        dictionary.Items.Remove("EmbeddedFiles");
        dictionary.Items.Remove("AF");
        dictionary.Items.Remove("EF");

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

            PdfObject? resolved = Resolve(objects, item);
            if (resolved is PdfDictionary action && TryGetForbiddenAction(objects, action, policy, out _, out _)) {
                dictionary.Items.Remove(key);
                continue;
            }

            if (key == "Next" && resolved is PdfArray nextActions) {
                FilterActions(objects, nextActions, policy);
            }

            if (key == "URI" && resolved is PdfDictionary uriDictionary &&
                TryGetString(objects, uriDictionary, "Base", out string? baseUri) && !policy.IsUriAllowed(baseUri!)) {
                uriDictionary.Items.Remove("Base");
            }

            if (item is not PdfReference) {
                SanitizeObject(objects, item, policy);
            }
        }
    }

    private static void FilterActions(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray actions,
        PdfSanitizationOptions policy) {
        for (int i = actions.Items.Count - 1; i >= 0; i--) {
            if (Resolve(objects, actions.Items[i]) is PdfDictionary action &&
                TryGetForbiddenAction(objects, action, policy, out _, out _)) {
                actions.Items.RemoveAt(i);
            }
        }
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
        out string? detail) {
        kind = PdfSanitizationFindingKind.ActiveAction;
        detail = null;
        if (Resolve(objects, dictionary.Get<PdfObject>("S")) is not PdfName actionName) {
            return false;
        }

        string actionType = actionName.Name;
        if (actionType == "URI") {
            if (TryGetString(objects, dictionary, "URI", out string? uri) && !policy.IsUriAllowed(uri!)) {
                kind = PdfSanitizationFindingKind.UnsafeUri;
                detail = uri;
                return true;
            }

            return false;
        }

        if (!PdfActiveContentPolicy.IsUnsafeActionType(actionType) || policy.IsActionAllowed(actionType)) {
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
