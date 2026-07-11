namespace OfficeIMO.OpenDocument;

/// <summary>Bounds deterministic OpenFormula parsing, dependency traversal, and recalculation.</summary>
public sealed class OdsFormulaEvaluationOptions {
    /// <summary>Maximum parser and evaluator operations in one evaluation context.</summary>
    public int MaximumOperations { get; set; } = 100_000;
    /// <summary>Maximum characters in one formula expression.</summary>
    public int MaximumFormulaCharacters { get; set; } = 32_768;
    /// <summary>Maximum cells expanded across range arguments.</summary>
    public int MaximumRangeCells { get; set; } = 100_000;
    /// <summary>Maximum dependency recursion depth.</summary>
    public int MaximumDependencyDepth { get; set; } = 64;
    /// <summary>Maximum formula cells updated by one recalculation.</summary>
    public int MaximumFormulaCells { get; set; } = 10_000;

    internal OdsFormulaEvaluationOptions Normalize() => new OdsFormulaEvaluationOptions {
        MaximumOperations = Math.Max(1, MaximumOperations),
        MaximumFormulaCharacters = Math.Max(1, MaximumFormulaCharacters),
        MaximumRangeCells = Math.Max(1, MaximumRangeCells),
        MaximumDependencyDepth = Math.Max(1, MaximumDependencyDepth),
        MaximumFormulaCells = Math.Max(1, MaximumFormulaCells)
    };
}

/// <summary>Result of evaluating one expression or formula cell.</summary>
public sealed class OdsFormulaEvaluationResult {
    internal OdsFormulaEvaluationResult(OdsFormulaValue value, int operations) { Value = value; Operations = operations; }
    /// <summary>Scalar result or an error value.</summary>
    public OdsFormulaValue Value { get; }
    /// <summary>Whether evaluation completed without an error value.</summary>
    public bool Success => Value.Kind != OdsFormulaValueKind.Error;
    /// <summary>Number of bounded evaluator operations consumed.</summary>
    public int Operations { get; }
    /// <summary>Error text when evaluation failed.</summary>
    public string? Error => Success ? null : Value.AsText();
}

/// <summary>One formula-cell recalculation failure.</summary>
public sealed class OdsFormulaDiagnostic {
    internal OdsFormulaDiagnostic(string sheetName, long row, long column, string message) {
        SheetName = sheetName; Row = row; Column = column; Message = message;
    }
    /// <summary>Worksheet name.</summary>
    public string SheetName { get; }
    /// <summary>Zero-based row.</summary>
    public long Row { get; }
    /// <summary>Zero-based column.</summary>
    public long Column { get; }
    /// <summary>Failure detail.</summary>
    public string Message { get; }
}

/// <summary>Bounded recalculation report.</summary>
public sealed class OdsRecalculationReport {
    private readonly List<OdsFormulaDiagnostic> _diagnostics = new List<OdsFormulaDiagnostic>();
    /// <summary>Formula cells discovered within the configured bound.</summary>
    public int FormulaCells { get; internal set; }
    /// <summary>Formula cells whose cached values were updated.</summary>
    public int UpdatedCells { get; internal set; }
    /// <summary>Formula cells that failed evaluation.</summary>
    public int FailedCells => _diagnostics.Count;
    /// <summary>Whether further formula cells were skipped because a bound was reached.</summary>
    public bool Truncated { get; internal set; }
    /// <summary>Cell-specific failures.</summary>
    public IReadOnlyList<OdsFormulaDiagnostic> Diagnostics => _diagnostics;
    internal void AddFailure(string sheetName, long row, long column, string message) =>
        _diagnostics.Add(new OdsFormulaDiagnostic(sheetName, row, column, message));
}
