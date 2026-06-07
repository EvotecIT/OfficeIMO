namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    private IReadOnlyList<string>? _catalogActionNames;
    private IReadOnlyList<string>? _catalogActionTypes;
    private IReadOnlyList<string>? _catalogActionSources;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfCatalogAction>>? _catalogActionsByActionType;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfCatalogAction>>? _catalogActionsBySource;

    /// <summary>Catalog-level actions read from supported name trees.</summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions { get; }

    /// <summary>Number of catalog-level actions read from supported name trees.</summary>
    public int CatalogActionCount => CatalogActions.Count;

    /// <summary>True when at least one catalog-level action was read from supported name trees.</summary>
    public bool HasCatalogActions => CatalogActionCount > 0;

    /// <summary>Catalog action name-tree keys in first-seen order.</summary>
    public IReadOnlyList<string> CatalogActionNames {
        get {
            if (_catalogActionNames is not null) {
                return _catalogActionNames;
            }

            var names = new List<string>(CatalogActions.Count);
            for (int i = 0; i < CatalogActions.Count; i++) {
                names.Add(CatalogActions[i].Name);
            }

            _catalogActionNames = names.AsReadOnly();
            return _catalogActionNames;
        }
    }

    /// <summary>Distinct catalog action types in first-seen order.</summary>
    public IReadOnlyList<string> CatalogActionTypes {
        get {
            if (_catalogActionTypes is not null) {
                return _catalogActionTypes;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var actionTypes = new List<string>();
            for (int i = 0; i < CatalogActions.Count; i++) {
                string actionType = CatalogActions[i].ActionType;
                if (seen.Add(actionType)) {
                    actionTypes.Add(actionType);
                }
            }

            _catalogActionTypes = actionTypes.AsReadOnly();
            return _catalogActionTypes;
        }
    }

    /// <summary>Distinct catalog action sources in first-seen order.</summary>
    public IReadOnlyList<string> CatalogActionSources {
        get {
            if (_catalogActionSources is not null) {
                return _catalogActionSources;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var sources = new List<string>();
            for (int i = 0; i < CatalogActions.Count; i++) {
                string source = CatalogActions[i].Source;
                if (seen.Add(source)) {
                    sources.Add(source);
                }
            }

            _catalogActionSources = sources.AsReadOnly();
            return _catalogActionSources;
        }
    }

    /// <summary>Catalog actions grouped by PDF action type.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfCatalogAction>> CatalogActionsByActionType {
        get {
            if (_catalogActionsByActionType is not null) {
                return _catalogActionsByActionType;
            }

            var grouped = new Dictionary<string, List<PdfCatalogAction>>(StringComparer.Ordinal);
            for (int i = 0; i < CatalogActions.Count; i++) {
                PdfCatalogAction action = CatalogActions[i];
                if (!grouped.TryGetValue(action.ActionType, out List<PdfCatalogAction>? actions)) {
                    actions = new List<PdfCatalogAction>();
                    grouped.Add(action.ActionType, actions);
                }

                actions.Add(action);
            }

            _catalogActionsByActionType = ToReadOnlyLookup(grouped);
            return _catalogActionsByActionType;
        }
    }

    /// <summary>Catalog actions grouped by catalog source.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfCatalogAction>> CatalogActionsBySource {
        get {
            if (_catalogActionsBySource is not null) {
                return _catalogActionsBySource;
            }

            var grouped = new Dictionary<string, List<PdfCatalogAction>>(StringComparer.Ordinal);
            for (int i = 0; i < CatalogActions.Count; i++) {
                PdfCatalogAction action = CatalogActions[i];
                if (!grouped.TryGetValue(action.Source, out List<PdfCatalogAction>? actions)) {
                    actions = new List<PdfCatalogAction>();
                    grouped.Add(action.Source, actions);
                }

                actions.Add(action);
            }

            _catalogActionsBySource = ToReadOnlyLookup(grouped);
            return _catalogActionsBySource;
        }
    }

    /// <summary>Returns catalog-level actions with a matching PDF action type.</summary>
    public IReadOnlyList<PdfCatalogAction> GetCatalogActionsByActionType(string actionType) {
        Guard.NotNullOrWhiteSpace(actionType, nameof(actionType));
        return CatalogActionsByActionType.TryGetValue(actionType, out IReadOnlyList<PdfCatalogAction>? actions)
            ? actions
            : Array.Empty<PdfCatalogAction>();
    }

    /// <summary>Returns catalog-level actions from a matching catalog source.</summary>
    public IReadOnlyList<PdfCatalogAction> GetCatalogActionsBySource(string source) {
        Guard.NotNullOrWhiteSpace(source, nameof(source));
        return CatalogActionsBySource.TryGetValue(source, out IReadOnlyList<PdfCatalogAction>? actions)
            ? actions
            : Array.Empty<PdfCatalogAction>();
    }
}
