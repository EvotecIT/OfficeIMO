namespace OfficeIMO.OpenDocument;

/// <summary>Bounded, side-effect-free evaluator for a documented OpenFormula subset.</summary>
public static class OdsFormulaEvaluator {
    /// <summary>Evaluates a formula cell without changing its cached value.</summary>
    public static OdsFormulaEvaluationResult EvaluateCell(OdsDocument document, string sheetName, long row, long column,
        OdsFormulaEvaluationOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        var context = new OdsFormulaEvaluationContext(document, options);
        OdsFormulaValue value = EvaluateCell(context, sheetName, row, column, 0);
        return new OdsFormulaEvaluationResult(value, context.Operations);
    }

    /// <summary>Evaluates an OpenFormula expression in the context of a worksheet.</summary>
    public static OdsFormulaEvaluationResult EvaluateExpression(OdsDocument document, string sheetName, string formula,
        OdsFormulaEvaluationOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));
        if (formula == null) throw new ArgumentNullException(nameof(formula));
        var context = new OdsFormulaEvaluationContext(document, options);
        OdsFormulaValue value;
        try {
            value = new OdsFormulaParser(formula, context, sheetName, 0).Parse();
        } catch (OdsFormulaException ex) { value = OdsFormulaValue.Error(ex.Message); }
        catch (Exception ex) when (ex is FormatException || ex is OverflowException || ex is ArgumentOutOfRangeException) {
            value = OdsFormulaValue.Error("Invalid formula value or reference: " + ex.Message);
        }
        return new OdsFormulaEvaluationResult(value, context.Operations);
    }

    internal static OdsFormulaValue EvaluateCell(OdsFormulaEvaluationContext context, string sheetName, long row, long column, int depth) {
        try {
            context.Step();
            if (depth > context.Options.MaximumDependencyDepth) throw new OdsFormulaException("Formula dependency depth limit exceeded.");
            var key = new OdsFormulaCellKey(sheetName, row, column);
            if (context.TryGetMemo(key, out OdsFormulaValue memo)) return memo;
            if (!context.Enter(key)) throw new OdsFormulaException("Cyclic formula dependency detected at " + key + ".");
            try {
                OdsSheet sheet = context.Document.GetSheet(sheetName) ?? throw new OdsFormulaException("Worksheet '" + sheetName + "' does not exist.");
                string? formula = sheet.GetFormula(row, column);
                OdsFormulaValue value = formula == null
                    ? OdsFormulaValue.FromCellValue(sheet.GetValue(row, column))
                    : new OdsFormulaParser(formula, context, sheetName, depth + 1).Parse();
                context.Memoize(key, value);
                return value;
            } finally { context.Exit(key); }
        } catch (OdsFormulaException ex) { return OdsFormulaValue.Error(ex.Message); }
        catch (Exception ex) when (ex is FormatException || ex is OverflowException || ex is ArgumentOutOfRangeException) {
            return OdsFormulaValue.Error("Invalid formula value or reference: " + ex.Message);
        }
    }
}

internal sealed class OdsFormulaEvaluationContext {
    private readonly Dictionary<OdsFormulaCellKey, OdsFormulaValue> _memo = new Dictionary<OdsFormulaCellKey, OdsFormulaValue>();
    private readonly HashSet<OdsFormulaCellKey> _visiting = new HashSet<OdsFormulaCellKey>();
    private int _rangeCells;
    internal OdsFormulaEvaluationContext(OdsDocument document, OdsFormulaEvaluationOptions? options) {
        Document = document;
        Options = (options ?? new OdsFormulaEvaluationOptions()).Normalize();
    }
    internal OdsDocument Document { get; }
    internal OdsFormulaEvaluationOptions Options { get; }
    internal int Operations { get; private set; }
    internal void Step() {
        Operations++;
        if (Operations > Options.MaximumOperations) throw new OdsFormulaException("Formula operation limit exceeded.");
    }
    internal void AddRangeCell() {
        _rangeCells++;
        if (_rangeCells > Options.MaximumRangeCells) throw new OdsFormulaException("Formula range-cell limit exceeded.");
    }
    internal bool TryGetMemo(OdsFormulaCellKey key, out OdsFormulaValue value) => _memo.TryGetValue(key, out value);
    internal void Memoize(OdsFormulaCellKey key, OdsFormulaValue value) => _memo[key] = value;
    internal bool Enter(OdsFormulaCellKey key) => _visiting.Add(key);
    internal void Exit(OdsFormulaCellKey key) => _visiting.Remove(key);
}

internal readonly struct OdsFormulaCellKey : IEquatable<OdsFormulaCellKey> {
    internal OdsFormulaCellKey(string sheetName, long row, long column) { SheetName = sheetName; Row = row; Column = column; }
    internal string SheetName { get; }
    internal long Row { get; }
    internal long Column { get; }
    public bool Equals(OdsFormulaCellKey other) => Row == other.Row && Column == other.Column &&
        string.Equals(SheetName, other.SheetName, StringComparison.Ordinal);
    public override bool Equals(object? obj) => obj is OdsFormulaCellKey other && Equals(other);
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(SheetName) ^ Row.GetHashCode() ^ Column.GetHashCode();
    public override string ToString() => SheetName + "!R" + (Row + 1).ToString(CultureInfo.InvariantCulture) + "C" + (Column + 1).ToString(CultureInfo.InvariantCulture);
}

internal sealed class OdsFormulaException : Exception {
    internal OdsFormulaException(string message) : base(message) { }
}
