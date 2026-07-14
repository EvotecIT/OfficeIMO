namespace OfficeIMO.Pdf;

/// <summary>Read-only document portfolio metadata discovered from the catalog collection dictionary.</summary>
public sealed class PdfPortfolioInfo {
    internal PdfPortfolioInfo(
        string? view,
        string? initialDocumentFileName,
        IReadOnlyList<PdfPortfolioFieldInfo> fields,
        string? sortField,
        bool? sortAscending) {
        View = view;
        InitialDocumentFileName = initialDocumentFileName;
        Fields = fields;
        SortField = sortField;
        SortAscending = sortAscending;
    }

    /// <summary>Raw collection view name: D, T, or H.</summary>
    public string? View { get; }
    /// <summary>Initial embedded document file name, when configured.</summary>
    public string? InitialDocumentFileName { get; }
    /// <summary>Portfolio schema fields.</summary>
    public IReadOnlyList<PdfPortfolioFieldInfo> Fields { get; }
    /// <summary>Portfolio schema key used for sorting, when configured.</summary>
    public string? SortField { get; }
    /// <summary>Sort direction, when configured.</summary>
    public bool? SortAscending { get; }
}

/// <summary>Read-only portfolio schema field metadata.</summary>
public sealed class PdfPortfolioFieldInfo {
    internal PdfPortfolioFieldInfo(string key, string? displayName, string? subtype, int? order, bool? visible, bool? editable) {
        Key = key;
        DisplayName = displayName;
        Subtype = subtype;
        Order = order;
        Visible = visible;
        Editable = editable;
    }

    /// <summary>Collection schema key.</summary>
    public string Key { get; }
    /// <summary>Displayed field name.</summary>
    public string? DisplayName { get; }
    /// <summary>Raw PDF collection-field subtype.</summary>
    public string? Subtype { get; }
    /// <summary>Display order.</summary>
    public int? Order { get; }
    /// <summary>Viewer visibility flag.</summary>
    public bool? Visible { get; }
    /// <summary>Viewer editability flag.</summary>
    public bool? Editable { get; }
}
