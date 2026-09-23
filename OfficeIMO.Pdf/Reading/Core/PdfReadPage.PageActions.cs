using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>Reads page-level additional action metadata from this page dictionary.</summary>
    public IReadOnlyList<PdfPageAction> GetPageActions() {
        _demandContentExtraction?.Invoke("page action");
        return GetPageActionsUnchecked();
    }

    internal IReadOnlyList<PdfPageAction> GetPageActionsUnchecked(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return _pageDict.Items.TryGetValue("AA", out var additionalActionsObject)
            ? ReadPageActions(additionalActionsObject, cancellationToken)
            : Array.Empty<PdfPageAction>();
    }

    private IReadOnlyList<PdfPageAction> ReadPageActions(PdfObject? obj, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var additionalActions = ResolveDictionary(obj);
        if (additionalActions is null || additionalActions.Items.Count == 0) {
            return Array.Empty<PdfPageAction>();
        }

        var actions = new List<PdfPageAction>();
        foreach (var item in additionalActions.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(item.Key)) {
                continue;
            }

            AddPageAction(item.Key, item.Key, item.Value, actions, new HashSet<int>(), cancellationToken);
        }

        return actions.Count == 0 ? Array.Empty<PdfPageAction>() : actions.AsReadOnly();
    }

    private void AddPageAction(
        string triggerName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfPageAction> result,
        HashSet<int> visitedReferences,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        bool enteredReference = TryEnterActionReference(actionObject, visitedReferences);
        if (!enteredReference) {
            return;
        }

        try {
            PdfObject? resolved = ResolveObject(actionObject);
            if (resolved is not PdfDictionary dictionary) {
                return;
            }

            string? actionType = TryReadActionType(dictionary);
            if (!string.IsNullOrEmpty(actionType)) {
                string? uri = null;
                if (string.Equals(actionType, "URI", StringComparison.Ordinal) &&
                    dictionary.Items.TryGetValue("URI", out PdfObject? uriObject) &&
                    ResolveObject(uriObject) is PdfStringObj uriText) {
                    uri = uriText.Value;
                }
                result.Add(new PdfPageAction(null, triggerName, actionType!, actionPath, uri, PdfActionPayloadFingerprint.Create(dictionary, _objects, _limits, cancellationToken)));
            }

            if (dictionary.Items.TryGetValue("Next", out var nextAction)) {
                AddPageNextActions(triggerName, actionPath + ".Next", nextAction, result, visitedReferences, cancellationToken);
            }
        } finally {
            LeaveActionReference(actionObject, visitedReferences);
        }
    }

    private void AddPageNextActions(
        string triggerName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfPageAction> result,
        HashSet<int> visitedReferences,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        bool enteredReference = TryEnterActionReference(actionObject, visitedReferences);
        if (!enteredReference) {
            return;
        }

        try {
            PdfObject? resolved = ResolveObject(actionObject);
            if (resolved is PdfArray array) {
                int activeIndex = 0;
                for (int i = 0; i < array.Items.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int before = result.Count;
                    AddPageAction(triggerName, actionPath + "." + activeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), array.Items[i], result, visitedReferences, cancellationToken);
                    if (result.Count > before) {
                        activeIndex++;
                    }
                }

                return;
            }

            AddPageAction(triggerName, actionPath, resolved, result, visitedReferences, cancellationToken);
        } finally {
            LeaveActionReference(actionObject, visitedReferences);
        }
    }

    private static bool TryEnterActionReference(PdfObject? actionObject, HashSet<int> visitedReferences) {
        return actionObject is not PdfReference reference || visitedReferences.Add(reference.ObjectNumber);
    }

    private static void LeaveActionReference(PdfObject? actionObject, HashSet<int> visitedReferences) {
        if (actionObject is PdfReference reference) {
            visitedReferences.Remove(reference.ObjectNumber);
        }
    }
}
