namespace OfficeIMO.OpenDocument;

/// <summary>Scalar value kinds produced by the bounded OpenFormula evaluator.</summary>
public enum OdsFormulaValueKind {
    /// <summary>No value.</summary>
    Empty,
    /// <summary>Numeric value.</summary>
    Number,
    /// <summary>Text value.</summary>
    Text,
    /// <summary>Boolean value.</summary>
    Boolean,
    /// <summary>Evaluation error.</summary>
    Error
}

/// <summary>Immutable scalar formula result.</summary>
public readonly struct OdsFormulaValue {
    private readonly double _number;
    private readonly string? _text;
    private readonly bool _boolean;

    private OdsFormulaValue(OdsFormulaValueKind kind, double number, string? text, bool boolean) {
        Kind = kind; _number = number; _text = text; _boolean = boolean;
    }

    /// <summary>Empty formula value.</summary>
    public static OdsFormulaValue Empty => new OdsFormulaValue(OdsFormulaValueKind.Empty, 0D, null, false);
    /// <summary>Creates a numeric value.</summary>
    public static OdsFormulaValue Number(double value) => new OdsFormulaValue(OdsFormulaValueKind.Number, value, null, false);
    /// <summary>Creates a text value.</summary>
    public static OdsFormulaValue Text(string? value) => new OdsFormulaValue(OdsFormulaValueKind.Text, 0D, value ?? string.Empty, false);
    /// <summary>Creates a Boolean value.</summary>
    public static OdsFormulaValue Boolean(bool value) => new OdsFormulaValue(OdsFormulaValueKind.Boolean, 0D, null, value);
    /// <summary>Creates an error value.</summary>
    public static OdsFormulaValue Error(string message) => new OdsFormulaValue(OdsFormulaValueKind.Error, 0D, message ?? "Formula evaluation failed.", false);

    /// <summary>Result kind.</summary>
    public OdsFormulaValueKind Kind { get; }
    /// <summary>Numeric value.</summary>
    public double AsNumber() {
        if (Kind == OdsFormulaValueKind.Number) return _number;
        if (Kind == OdsFormulaValueKind.Boolean) return _boolean ? 1D : 0D;
        if (Kind == OdsFormulaValueKind.Empty) return 0D;
        if (Kind == OdsFormulaValueKind.Text && double.TryParse(_text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) return value;
        throw new InvalidOperationException("Formula value is not numeric.");
    }
    /// <summary>Boolean value.</summary>
    public bool AsBoolean() {
        if (Kind == OdsFormulaValueKind.Boolean) return _boolean;
        if (Kind == OdsFormulaValueKind.Number) return _number != 0D;
        if (Kind == OdsFormulaValueKind.Empty) return false;
        if (Kind == OdsFormulaValueKind.Text && bool.TryParse(_text, out bool value)) return value;
        throw new InvalidOperationException("Formula value is not Boolean.");
    }
    /// <summary>Text or error payload.</summary>
    public string AsText() {
        switch (Kind) {
            case OdsFormulaValueKind.Empty: return string.Empty;
            case OdsFormulaValueKind.Number: return _number.ToString("R", CultureInfo.InvariantCulture);
            case OdsFormulaValueKind.Boolean: return _boolean ? "TRUE" : "FALSE";
            default: return _text ?? string.Empty;
        }
    }
    /// <inheritdoc />
    public override string ToString() => AsText();

    internal static OdsFormulaValue FromCellValue(OdsCellValue value) {
        try {
            switch (value.Kind) {
                case OdsCellValueKind.Empty: return Empty;
                case OdsCellValueKind.String: return Text(value.LexicalValue);
                case OdsCellValueKind.Boolean: return Boolean(value.AsBoolean());
                case OdsCellValueKind.Number:
                case OdsCellValueKind.Percentage:
                case OdsCellValueKind.Currency: return Number(value.AsDouble());
                case OdsCellValueKind.Date: return Text(value.LexicalValue);
                case OdsCellValueKind.Time: return Number(value.AsTimeSpan().TotalDays);
                default: return Text(value.ToString());
            }
        } catch (FormatException) { return Error("The referenced cell has an invalid typed value."); }
        catch (OverflowException) { return Error("The referenced cell value is outside the evaluator's numeric range."); }
    }
}
