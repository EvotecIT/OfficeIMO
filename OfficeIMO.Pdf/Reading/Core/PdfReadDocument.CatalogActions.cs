namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfCatalogAction> ExtractCatalogActions() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            return Array.Empty<PdfCatalogAction>();
        }

        var result = new List<PdfCatalogAction>();
        if (catalog.Items.TryGetValue("Names", out var namesObject) &&
            ResolveDict(namesObject) is PdfDictionary namesDictionary &&
            namesDictionary.Items.TryGetValue("JavaScript", out var javaScriptNameTree)) {
            AddCatalogActionsFromNameTree(javaScriptNameTree, result, new HashSet<int>());
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

        return result.Count == 0 ? Array.Empty<PdfCatalogAction>() : result.AsReadOnly();
    }

    private void AddCatalogActionsFromNameTree(
        PdfObject treeObject,
        List<PdfCatalogAction> result,
        HashSet<int> visitedReferences) {
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(_objects, reference, out var indirect)) {
                return;
            }

            AddCatalogActionsFromNameTree(indirect.Value, result, visitedReferences);
            return;
        }

        if (treeObject is not PdfDictionary tree) {
            return;
        }

        if (tree.Items.TryGetValue("Names", out var actionNamesObject) &&
            ResolveArray(actionNamesObject) is PdfArray actionNames) {
            for (int i = 0; i + 1 < actionNames.Items.Count; i += 2) {
                if (TryReadCatalogActionName(actionNames.Items[i], out string? name)) {
                    AddCatalogAction(name!, "Names/JavaScript", null, actionNames.Items[i + 1], result, new HashSet<int>());
                }
            }
        }

        if (tree.Items.TryGetValue("Kids", out var kidsObject) &&
            ResolveArray(kidsObject) is PdfArray kids) {
            foreach (var kid in kids.Items) {
                AddCatalogActionsFromNameTree(kid, result, visitedReferences);
            }
        }
    }

    private bool TryReadCatalogActionName(PdfObject obj, out string? name) {
        switch (ResolveObject(obj)) {
            case PdfStringObj text:
                name = text.Value;
                return !string.IsNullOrEmpty(name);
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
        HashSet<int> visitedReferences) {
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference && !visitedReferences.Add(reference.ObjectNumber)) {
            return;
        }

        if (resolved is not PdfDictionary dictionary) {
            return;
        }

        if (TryReadCatalogActionType(dictionary, out string? actionType) &&
            IsActiveCatalogActionType(actionType!)) {
            result.Add(new PdfCatalogAction(name, actionType!, source, triggerName));
        }

        if (dictionary.Items.TryGetValue("Next", out var nextAction)) {
            AddCatalogNextActions(name + ".Next", source, triggerName, nextAction, result, visitedReferences);
        }
    }

    private void AddCatalogNextActions(
        string name,
        string source,
        string? triggerName,
        PdfObject obj,
        List<PdfCatalogAction> result,
        HashSet<int> visitedReferences) {
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference && !visitedReferences.Add(reference.ObjectNumber)) {
            return;
        }

        if (resolved is PdfArray actions) {
            int activeIndex = 0;
            for (int i = 0; i < actions.Items.Count; i++) {
                int before = result.Count;
                AddCatalogAction(name + "." + activeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), source, triggerName, actions.Items[i], result, visitedReferences);
                if (result.Count > before) {
                    activeIndex++;
                }
            }

            return;
        }

        if (resolved is PdfDictionary) {
            AddCatalogAction(name, source, triggerName, resolved, result, visitedReferences);
        }
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
            case "RichMedia":
                return true;
            default:
                return false;
        }
    }
}
