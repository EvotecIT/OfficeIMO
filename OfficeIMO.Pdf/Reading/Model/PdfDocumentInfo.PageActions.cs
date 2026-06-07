namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    private IReadOnlyList<PdfPageAction>? _pageActions;
    private IReadOnlyList<string>? _pageActionTypes;
    private IReadOnlyList<string>? _pageActionTriggerNames;
    private IReadOnlyList<string>? _pageActionPaths;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>>? _pageActionsByActionType;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>>? _pageActionsByTriggerName;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>>? _pageActionsByActionPath;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfPageAction>>? _pageActionsByPageNumber;

    /// <summary>Number of page-level additional actions read from page dictionaries.</summary>
    public int PageActionCount => PageActions.Count;

    /// <summary>True when at least one page-level additional action was read from page dictionaries.</summary>
    public bool HasPageActions => PageActionCount > 0;

    /// <summary>Page-level additional actions read from page dictionaries in document order.</summary>
    public IReadOnlyList<PdfPageAction> PageActions {
        get {
            if (_pageActions is not null) {
                return _pageActions;
            }

            var actions = new List<PdfPageAction>();
            for (int i = 0; i < Pages.Count; i++) {
                PdfPageInfo page = Pages[i];
                for (int j = 0; j < page.PageActions.Count; j++) {
                    PdfPageAction action = page.PageActions[j];
                    actions.Add(action.PageNumber.HasValue ? action : action.WithPageNumber(page.PageNumber));
                }
            }

            _pageActions = actions.AsReadOnly();
            return _pageActions;
        }
    }

    /// <summary>Distinct page-level action types in first-seen document order.</summary>
    public IReadOnlyList<string> PageActionTypes {
        get {
            if (_pageActionTypes is not null) {
                return _pageActionTypes;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var actionTypes = new List<string>();
            for (int i = 0; i < PageActions.Count; i++) {
                string actionType = PageActions[i].ActionType;
                if (seen.Add(actionType)) {
                    actionTypes.Add(actionType);
                }
            }

            _pageActionTypes = actionTypes.AsReadOnly();
            return _pageActionTypes;
        }
    }

    /// <summary>Distinct page-level additional-action trigger keys in first-seen document order.</summary>
    public IReadOnlyList<string> PageActionTriggerNames {
        get {
            if (_pageActionTriggerNames is not null) {
                return _pageActionTriggerNames;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var triggerNames = new List<string>();
            for (int i = 0; i < PageActions.Count; i++) {
                string triggerName = PageActions[i].TriggerName;
                if (seen.Add(triggerName)) {
                    triggerNames.Add(triggerName);
                }
            }

            _pageActionTriggerNames = triggerNames.AsReadOnly();
            return _pageActionTriggerNames;
        }
    }

    /// <summary>Distinct page-level action paths in first-seen document order.</summary>
    public IReadOnlyList<string> PageActionPaths {
        get {
            if (_pageActionPaths is not null) {
                return _pageActionPaths;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var actionPaths = new List<string>();
            for (int i = 0; i < PageActions.Count; i++) {
                string actionPath = PageActions[i].ActionPath;
                if (seen.Add(actionPath)) {
                    actionPaths.Add(actionPath);
                }
            }

            _pageActionPaths = actionPaths.AsReadOnly();
            return _pageActionPaths;
        }
    }

    /// <summary>Page-level additional actions grouped by PDF action type.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>> PageActionsByActionType {
        get {
            if (_pageActionsByActionType is not null) {
                return _pageActionsByActionType;
            }

            var grouped = new Dictionary<string, List<PdfPageAction>>(StringComparer.Ordinal);
            for (int i = 0; i < PageActions.Count; i++) {
                AddPageAction(grouped, PageActions[i].ActionType, PageActions[i]);
            }

            _pageActionsByActionType = ToReadOnlyLookup(grouped);
            return _pageActionsByActionType;
        }
    }

    /// <summary>Page-level additional actions grouped by page /AA trigger key.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>> PageActionsByTriggerName {
        get {
            if (_pageActionsByTriggerName is not null) {
                return _pageActionsByTriggerName;
            }

            var grouped = new Dictionary<string, List<PdfPageAction>>(StringComparer.Ordinal);
            for (int i = 0; i < PageActions.Count; i++) {
                AddPageAction(grouped, PageActions[i].TriggerName, PageActions[i]);
            }

            _pageActionsByTriggerName = ToReadOnlyLookup(grouped);
            return _pageActionsByTriggerName;
        }
    }

    /// <summary>Page-level additional actions grouped by stable action path.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfPageAction>> PageActionsByActionPath {
        get {
            if (_pageActionsByActionPath is not null) {
                return _pageActionsByActionPath;
            }

            var grouped = new Dictionary<string, List<PdfPageAction>>(StringComparer.Ordinal);
            for (int i = 0; i < PageActions.Count; i++) {
                AddPageAction(grouped, PageActions[i].ActionPath, PageActions[i]);
            }

            _pageActionsByActionPath = ToReadOnlyLookup(grouped);
            return _pageActionsByActionPath;
        }
    }

    /// <summary>Page-level additional actions grouped by one-based page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfPageAction>> PageActionsByPageNumber {
        get {
            if (_pageActionsByPageNumber is not null) {
                return _pageActionsByPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfPageAction>>();
            for (int i = 0; i < PageActions.Count; i++) {
                PdfPageAction action = PageActions[i];
                if (action.PageNumber.HasValue) {
                    if (!grouped.TryGetValue(action.PageNumber.Value, out List<PdfPageAction>? actions)) {
                        actions = new List<PdfPageAction>();
                        grouped.Add(action.PageNumber.Value, actions);
                    }

                    actions.Add(action);
                }
            }

            _pageActionsByPageNumber = ToReadOnlyLookup(grouped);
            return _pageActionsByPageNumber;
        }
    }

    /// <summary>Returns page-level additional actions with a matching PDF action type.</summary>
    public IReadOnlyList<PdfPageAction> GetPageActionsByActionType(string actionType) {
        Guard.NotNullOrWhiteSpace(actionType, nameof(actionType));
        return PageActionsByActionType.TryGetValue(actionType, out IReadOnlyList<PdfPageAction>? actions)
            ? actions
            : Array.Empty<PdfPageAction>();
    }

    /// <summary>Returns page-level additional actions with a matching page /AA trigger key.</summary>
    public IReadOnlyList<PdfPageAction> GetPageActionsByTriggerName(string triggerName) {
        Guard.NotNullOrWhiteSpace(triggerName, nameof(triggerName));
        return PageActionsByTriggerName.TryGetValue(triggerName, out IReadOnlyList<PdfPageAction>? actions)
            ? actions
            : Array.Empty<PdfPageAction>();
    }

    /// <summary>Returns page-level additional actions with a matching stable action path.</summary>
    public IReadOnlyList<PdfPageAction> GetPageActionsByActionPath(string actionPath) {
        Guard.NotNullOrWhiteSpace(actionPath, nameof(actionPath));
        return PageActionsByActionPath.TryGetValue(actionPath, out IReadOnlyList<PdfPageAction>? actions)
            ? actions
            : Array.Empty<PdfPageAction>();
    }

    /// <summary>Returns page-level additional actions for a one-based page number.</summary>
    public IReadOnlyList<PdfPageAction> GetPageActions(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return PageActionsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfPageAction>? actions)
            ? actions
            : Array.Empty<PdfPageAction>();
    }

    private static void AddPageAction(Dictionary<string, List<PdfPageAction>> grouped, string key, PdfPageAction action) {
        if (!grouped.TryGetValue(key, out List<PdfPageAction>? actions)) {
            actions = new List<PdfPageAction>();
            grouped.Add(key, actions);
        }

        actions.Add(action);
    }
}
