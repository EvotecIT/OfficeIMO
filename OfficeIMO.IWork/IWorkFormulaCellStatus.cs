namespace OfficeIMO.IWork;

/// <summary>The recovered source cache of a projected formula cell; it does not describe destination recalculation.</summary>
public enum IWorkFormulaCacheStatus {
    /// <summary>No cached value was recovered, including an unresolved cache reference.</summary>
    Missing,
    /// <summary>A cached value was recovered incompletely.</summary>
    Partial,
    /// <summary>A complete cached value was recovered. This does not establish freshness.</summary>
    Complete,
    /// <summary>A numeric cache exceeds portable precision, or an error cache was recovered as a generic marker without its original error code.</summary>
    Approximate,
    /// <summary>A supported header declares a formula, but its cell contents could not be decoded to assess the cache.</summary>
    Unassessed
}

/// <summary>Source expression and cache assessment for one cell known to declare a formula.</summary>
public sealed class IWorkFormulaCellStatus {
    internal IWorkFormulaCellStatus(IWorkTable table, IWorkTableCell cell) {
        TableIdentity = table.SourceIdentity;
        Row = cell.Row;
        Column = cell.Column;
        ExpressionIsAssessed = cell.Kind == IWorkCellKind.Formula;
        ExpressionIsComplete = cell.FormulaIsComplete;
        CacheStatus = !ExpressionIsAssessed ? IWorkFormulaCacheStatus.Unassessed
            : cell.Value == null ? IWorkFormulaCacheStatus.Missing
            : !cell.CachedValueIsComplete ? IWorkFormulaCacheStatus.Partial
            : cell.NumericValueIsApproximate || cell.ValueKind == IWorkCellKind.Error && cell.CachedDisplayText == "#ERROR" ? IWorkFormulaCacheStatus.Approximate
            : IWorkFormulaCacheStatus.Complete;
        CachedValueKind = cell.Value == null ? null : cell.ValueKind;
    }

    /// <summary>Gets the native table-info identity when available.</summary>
    public IWorkObjectIdentity? TableIdentity { get; }
    /// <summary>Gets the one-based source row.</summary>
    public int Row { get; }
    /// <summary>Gets the one-based source column.</summary>
    public int Column { get; }
    /// <summary>Gets whether expression reconstruction was assessed; false for a declared formula in an undecoded cell.</summary>
    public bool ExpressionIsAssessed { get; }
    /// <summary>Gets whether the projected expression is complete, independently of its cached value.</summary>
    public bool ExpressionIsComplete { get; }
    /// <summary>Gets the recovered cache assessment, or Unassessed when the cell contents could not be decoded.</summary>
    public IWorkFormulaCacheStatus CacheStatus { get; }
    /// <summary>Gets the recovered cache type, or null when no value was recovered.</summary>
    public IWorkCellKind? CachedValueKind { get; }
}

/// <summary>Source formula assessments over projected cells, including supported headers declaring formulas in undecoded cells. Unknown headers and inactive table records are excluded.</summary>
public sealed class IWorkFormulaSummary {
    internal IWorkFormulaSummary(IReadOnlyList<IWorkFormulaCellStatus> cells) {
        TotalCount = cells.Count;
        foreach (IWorkFormulaCellStatus cell in cells) {
            if (!cell.ExpressionIsAssessed) UnassessedExpressionCount++;
            else if (cell.ExpressionIsComplete) CompleteExpressionCount++;
            else IncompleteExpressionCount++;
            if (cell.CacheStatus == IWorkFormulaCacheStatus.Complete) CompleteCacheCount++;
            else if (cell.CacheStatus == IWorkFormulaCacheStatus.Partial) PartialCacheCount++;
            else if (cell.CacheStatus == IWorkFormulaCacheStatus.Approximate) ApproximateCacheCount++;
            else if (cell.CacheStatus == IWorkFormulaCacheStatus.Unassessed) UnassessedCacheCount++;
            else MissingCacheCount++;
        }
    }

    /// <summary>Gets the formula-cell count established by supported source headers, including undecoded cell contents.</summary>
    public int TotalCount { get; }
    /// <summary>Gets the count of complete projected expressions.</summary>
    public int CompleteExpressionCount { get; }
    /// <summary>Gets the count of incomplete or unresolved expressions.</summary>
    public int IncompleteExpressionCount { get; }
    /// <summary>Gets the count of declared formulas whose cell contents could not be decoded for expression assessment.</summary>
    public int UnassessedExpressionCount { get; }
    /// <summary>Gets the count with complete recovered caches, without assessing freshness.</summary>
    public int CompleteCacheCount { get; }
    /// <summary>Gets the count with partial recovered caches.</summary>
    public int PartialCacheCount { get; }
    /// <summary>Gets the count of numeric caches above portable precision and source error caches represented by generic markers.</summary>
    public int ApproximateCacheCount { get; }
    /// <summary>Gets the count without recovered caches.</summary>
    public int MissingCacheCount { get; }
    /// <summary>Gets the count of declared formulas whose cell contents could not be decoded for cache assessment.</summary>
    public int UnassessedCacheCount { get; }
}
