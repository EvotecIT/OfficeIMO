namespace OfficeIMO.IWork;

/// <summary>The recovered source cache of a projected formula cell; it does not describe destination recalculation.</summary>
public enum IWorkFormulaCacheStatus {
    /// <summary>No cached value was recovered, including an unresolved cache reference.</summary>
    Missing,
    /// <summary>A cached value was recovered incompletely.</summary>
    Partial,
    /// <summary>A complete cached value was recovered. This does not establish freshness.</summary>
    Complete,
    /// <summary>The source error state was recovered as a generic marker without its original error code.</summary>
    Approximate
}

/// <summary>Source expression and cache assessment for one projected formula cell.</summary>
public sealed class IWorkFormulaCellStatus {
    internal IWorkFormulaCellStatus(IWorkTable table, IWorkTableCell cell) {
        TableIdentity = table.SourceIdentity;
        Row = cell.Row;
        Column = cell.Column;
        ExpressionIsComplete = cell.FormulaIsComplete;
        CacheStatus = cell.Value == null ? IWorkFormulaCacheStatus.Missing
            : !cell.CachedValueIsComplete ? IWorkFormulaCacheStatus.Partial
            : cell.ValueKind == IWorkCellKind.Error && cell.CachedDisplayText == "#ERROR" ? IWorkFormulaCacheStatus.Approximate
            : IWorkFormulaCacheStatus.Complete;
        CachedValueKind = cell.Value == null ? null : cell.ValueKind;
    }

    /// <summary>Gets the native table-info identity when available.</summary>
    public IWorkObjectIdentity? TableIdentity { get; }
    /// <summary>Gets the one-based source row.</summary>
    public int Row { get; }
    /// <summary>Gets the one-based source column.</summary>
    public int Column { get; }
    /// <summary>Gets whether the projected expression is complete, independently of its cached value.</summary>
    public bool ExpressionIsComplete { get; }
    /// <summary>Gets whether a complete, partial, approximate, or no cached value was recovered.</summary>
    public IWorkFormulaCacheStatus CacheStatus { get; }
    /// <summary>Gets the recovered cache type, or null when no value was recovered.</summary>
    public IWorkCellKind? CachedValueKind { get; }
}

/// <summary>Source formula assessments over projected cells. Undecoded cells and inactive table records are excluded.</summary>
public sealed class IWorkFormulaSummary {
    internal IWorkFormulaSummary(IReadOnlyList<IWorkFormulaCellStatus> cells) {
        TotalCount = cells.Count;
        foreach (IWorkFormulaCellStatus cell in cells) {
            if (cell.ExpressionIsComplete) CompleteExpressionCount++;
            else IncompleteExpressionCount++;
            if (cell.CacheStatus == IWorkFormulaCacheStatus.Complete) CompleteCacheCount++;
            else if (cell.CacheStatus == IWorkFormulaCacheStatus.Partial) PartialCacheCount++;
            else if (cell.CacheStatus == IWorkFormulaCacheStatus.Approximate) ApproximateCacheCount++;
            else MissingCacheCount++;
        }
    }

    /// <summary>Gets the projected formula-cell count.</summary>
    public int TotalCount { get; }
    /// <summary>Gets the count of complete projected expressions.</summary>
    public int CompleteExpressionCount { get; }
    /// <summary>Gets the count of incomplete or unresolved expressions.</summary>
    public int IncompleteExpressionCount { get; }
    /// <summary>Gets the count with complete recovered caches, without assessing freshness.</summary>
    public int CompleteCacheCount { get; }
    /// <summary>Gets the count with partial recovered caches.</summary>
    public int PartialCacheCount { get; }
    /// <summary>Gets the count with generic error markers whose original codes were not decoded.</summary>
    public int ApproximateCacheCount { get; }
    /// <summary>Gets the count without recovered caches.</summary>
    public int MissingCacheCount { get; }
}
