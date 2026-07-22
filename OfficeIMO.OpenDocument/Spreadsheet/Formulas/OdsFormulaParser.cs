namespace OfficeIMO.OpenDocument;

internal sealed class OdsFormulaParser {
    private readonly OdsFormulaLexer _lexer;
    private readonly OdsFormulaEvaluationContext _context;
    private readonly string _sheetName;
    private readonly int _depth;
    private OdsFormulaToken _current;
    private int _syntaxDepth;

    internal OdsFormulaParser(string formula, OdsFormulaEvaluationContext context, string sheetName, int depth) {
        string expression = formula.Trim();
        int prefix = expression.IndexOf(":=", StringComparison.Ordinal);
        if (prefix >= 0) expression = expression.Substring(prefix + 2);
        else if (expression.StartsWith("=", StringComparison.Ordinal)) expression = expression.Substring(1);
        if (expression.Length > context.Options.MaximumFormulaCharacters) throw new OdsFormulaException("Formula character limit exceeded.");
        _lexer = new OdsFormulaLexer(expression);
        _context = context;
        _sheetName = sheetName;
        _depth = depth;
        _current = _lexer.Next();
    }

    internal OdsFormulaValue Parse() {
        OdsFormulaOperand operand = ParseComparison();
        if (_current.Kind != OdsFormulaTokenKind.End) throw Error("Unexpected token '" + _current.Text + "'.");
        return operand.RequireScalar();
    }

    private OdsFormulaOperand ParseComparison() {
        OdsFormulaOperand left = ParseConcat();
        while (IsComparison(_current.Kind)) {
            OdsFormulaTokenKind operation = Take().Kind;
            OdsFormulaOperand right = ParseConcat();
            left = Scalar(Compare(operation, left.RequireScalar(), right.RequireScalar()));
        }
        return left;
    }

    private OdsFormulaOperand ParseConcat() {
        OdsFormulaOperand left = ParseAdditive();
        while (_current.Kind == OdsFormulaTokenKind.Ampersand) {
            Take();
            OdsFormulaValue right = ParseAdditive().RequireScalar();
            OdsFormulaValue scalar = left.RequireScalar();
            left = Scalar(Propagate(scalar, right) ?? OdsFormulaValue.Text(scalar.AsText() + right.AsText()));
        }
        return left;
    }

    private OdsFormulaOperand ParseAdditive() {
        OdsFormulaOperand left = ParseMultiplicative();
        while (_current.Kind == OdsFormulaTokenKind.Plus || _current.Kind == OdsFormulaTokenKind.Minus) {
            OdsFormulaTokenKind operation = Take().Kind;
            OdsFormulaValue right = ParseMultiplicative().RequireScalar();
            left = Scalar(NumericBinary(operation, left.RequireScalar(), right));
        }
        return left;
    }

    private OdsFormulaOperand ParseMultiplicative() {
        OdsFormulaOperand left = ParseUnary();
        while (_current.Kind == OdsFormulaTokenKind.Star || _current.Kind == OdsFormulaTokenKind.Slash) {
            OdsFormulaTokenKind operation = Take().Kind;
            OdsFormulaValue right = ParseUnary().RequireScalar();
            left = Scalar(NumericBinary(operation, left.RequireScalar(), right));
        }
        return left;
    }

    private OdsFormulaOperand ParsePower() {
        OdsFormulaOperand left = ParsePostfix();
        if (_current.Kind == OdsFormulaTokenKind.Caret) {
            Take();
            EnterSyntax();
            try {
                OdsFormulaValue right = ParseUnary().RequireScalar();
                left = Scalar(NumericBinary(OdsFormulaTokenKind.Caret, left.RequireScalar(), right));
            } finally { _syntaxDepth--; }
        }
        return left;
    }

    private OdsFormulaOperand ParseUnary() {
        if (_current.Kind == OdsFormulaTokenKind.Plus || _current.Kind == OdsFormulaTokenKind.Minus) {
            bool negate = Take().Kind == OdsFormulaTokenKind.Minus;
            EnterSyntax();
            try {
                OdsFormulaOperand operand = ParseUnary();
                if (!negate) return operand;
                OdsFormulaValue value = operand.RequireScalar();
                return value.Kind == OdsFormulaValueKind.Error ? Scalar(value) : Scalar(Number(-RequireNumber(value)));
            } finally { _syntaxDepth--; }
        }
        return ParsePower();
    }

    private OdsFormulaOperand ParsePostfix() {
        OdsFormulaOperand result = ParsePrimary();
        while (_current.Kind == OdsFormulaTokenKind.Percent) {
            Take();
            OdsFormulaValue value = result.RequireScalar();
            result = value.Kind == OdsFormulaValueKind.Error ? Scalar(value) : Scalar(Number(RequireNumber(value) / 100D));
        }
        return result;
    }

    private OdsFormulaOperand ParsePrimary() {
        _context.Step();
        switch (_current.Kind) {
            case OdsFormulaTokenKind.Number:
                string lexical = Take().Text;
                if (!double.TryParse(lexical, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) throw Error("Invalid number '" + lexical + "'.");
                return Scalar(Number(number));
            case OdsFormulaTokenKind.String:
                return Scalar(OdsFormulaValue.Text(Take().Text));
            case OdsFormulaTokenKind.Reference:
                return ResolveReference(Take().Text);
            case OdsFormulaTokenKind.Identifier:
                return ParseIdentifier();
            case OdsFormulaTokenKind.LeftParenthesis:
                Take();
                EnterSyntax();
                try {
                    OdsFormulaOperand nested = ParseComparison();
                    Expect(OdsFormulaTokenKind.RightParenthesis);
                    return nested;
                } finally { _syntaxDepth--; }
            default:
                throw Error("Expected a value but found '" + _current.Text + "'.");
        }
    }

    private OdsFormulaOperand ParseIdentifier() {
        string name = Take().Text;
        if (_current.Kind != OdsFormulaTokenKind.LeftParenthesis) {
            if (string.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase)) return Scalar(OdsFormulaValue.Boolean(true));
            if (string.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase)) return Scalar(OdsFormulaValue.Boolean(false));
            throw Error("Unknown formula name '" + name + "'.");
        }
        Take();
        EnterSyntax();
        var arguments = new List<OdsFormulaOperand>();
        try {
            if (_current.Kind != OdsFormulaTokenKind.RightParenthesis) {
                while (true) {
                    arguments.Add(ParseComparison());
                    if (_current.Kind != OdsFormulaTokenKind.Separator) break;
                    Take();
                }
            }
            Expect(OdsFormulaTokenKind.RightParenthesis);
        } finally { _syntaxDepth--; }
        return Scalar(EvaluateFunction(name, arguments));
    }

    private OdsFormulaValue EvaluateFunction(string name, IReadOnlyList<OdsFormulaOperand> arguments) {
        _context.Step();
        string normalized = name.ToUpperInvariant();
        List<OdsFormulaValue> values = Flatten(arguments);
        OdsFormulaValue? error = values.FirstOrDefault(value => value.Kind == OdsFormulaValueKind.Error);
        if (error.HasValue && error.Value.Kind == OdsFormulaValueKind.Error) return error.Value;
        var numbers = new List<double>();
        foreach (OdsFormulaValue value in values) {
            if (value.Kind == OdsFormulaValueKind.Empty || value.Kind == OdsFormulaValueKind.Text) continue;
            numbers.Add(RequireNumber(value));
        }
        switch (normalized) {
            case "SUM": return Number(numbers.Sum());
            case "AVERAGE": return numbers.Count == 0 ? OdsFormulaValue.Error("AVERAGE requires at least one numeric value.") : Number(numbers.Average());
            case "MIN": return Number(numbers.Count == 0 ? 0D : numbers.Min());
            case "MAX": return Number(numbers.Count == 0 ? 0D : numbers.Max());
            case "COUNT": return Number(values.Count(value => value.Kind == OdsFormulaValueKind.Number));
            case "PRODUCT": return Number(numbers.Count == 0 ? 0D : numbers.Aggregate(1D, (left, right) => left * right));
            case "ABS": return UnaryFunction(normalized, arguments, value => Math.Abs(value));
            case "SQRT": return UnaryFunction(normalized, arguments, value => value < 0D ? double.NaN : Math.Sqrt(value));
            case "ROUND":
                if (arguments.Count < 1 || arguments.Count > 2) return OdsFormulaValue.Error("ROUND expects one or two scalar arguments.");
                int digits = arguments.Count == 2 ? checked((int)RequireNumber(arguments[1].RequireScalar())) : 0;
                digits = Math.Max(-15, Math.Min(15, digits));
                double roundValue = RequireNumber(arguments[0].RequireScalar());
                if (digits >= 0) return Number(Math.Round(roundValue, digits, MidpointRounding.AwayFromZero));
                double scale = Math.Pow(10D, -digits);
                return Number(Math.Round(roundValue / scale, 0, MidpointRounding.AwayFromZero) * scale);
            case "POWER":
                if (arguments.Count != 2) return OdsFormulaValue.Error("POWER expects two scalar arguments.");
                return Number(Math.Pow(RequireNumber(arguments[0].RequireScalar()), RequireNumber(arguments[1].RequireScalar())));
            default: return OdsFormulaValue.Error("Unsupported OpenFormula function '" + name + "'.");
        }
    }

    private OdsFormulaValue UnaryFunction(string name, IReadOnlyList<OdsFormulaOperand> arguments, Func<double, double> function) {
        if (arguments.Count != 1) return OdsFormulaValue.Error(name + " expects one scalar argument.");
        OdsFormulaValue value = arguments[0].RequireScalar();
        if (value.Kind == OdsFormulaValueKind.Error) return value;
        return Number(function(RequireNumber(value)));
    }

    private OdsFormulaOperand ResolveReference(string reference) {
        OdsFormulaReference range = OdsFormulaReference.Parse(reference, _sheetName);
        OdsSheet sheet = _context.Document.GetSheet(range.SheetName) ?? throw Error("Worksheet '" + range.SheetName + "' does not exist.");
        var values = new List<OdsFormulaValue>();
        for (long row = range.FirstRow; row <= range.LastRow; row++) {
            for (long column = range.FirstColumn; column <= range.LastColumn; column++) {
                _context.AddRangeCell();
                values.Add(OdsFormulaEvaluator.EvaluateCell(_context, sheet.Name, row, column, _depth + 1));
            }
        }
        return values.Count == 1 ? Scalar(values[0]) : OdsFormulaOperand.Range(values);
    }

    private OdsFormulaValue NumericBinary(OdsFormulaTokenKind operation, OdsFormulaValue left, OdsFormulaValue right) {
        OdsFormulaValue? error = Propagate(left, right);
        if (error.HasValue) return error.Value;
        double a = RequireNumber(left), b = RequireNumber(right), result;
        switch (operation) {
            case OdsFormulaTokenKind.Plus: result = a + b; break;
            case OdsFormulaTokenKind.Minus: result = a - b; break;
            case OdsFormulaTokenKind.Star: result = a * b; break;
            case OdsFormulaTokenKind.Slash:
                if (b == 0D) return OdsFormulaValue.Error("Division by zero.");
                result = a / b; break;
            default: result = Math.Pow(a, b); break;
        }
        return Number(result);
    }

    private OdsFormulaValue Compare(OdsFormulaTokenKind operation, OdsFormulaValue left, OdsFormulaValue right) {
        OdsFormulaValue? error = Propagate(left, right);
        if (error.HasValue) return error.Value;
        int comparison;
        if (IsNumericKind(left.Kind) && IsNumericKind(right.Kind)) comparison = RequireNumber(left).CompareTo(RequireNumber(right));
        else comparison = string.Compare(left.AsText(), right.AsText(), StringComparison.Ordinal);
        switch (operation) {
            case OdsFormulaTokenKind.Equal: return OdsFormulaValue.Boolean(comparison == 0);
            case OdsFormulaTokenKind.NotEqual: return OdsFormulaValue.Boolean(comparison != 0);
            case OdsFormulaTokenKind.Less: return OdsFormulaValue.Boolean(comparison < 0);
            case OdsFormulaTokenKind.LessOrEqual: return OdsFormulaValue.Boolean(comparison <= 0);
            case OdsFormulaTokenKind.Greater: return OdsFormulaValue.Boolean(comparison > 0);
            default: return OdsFormulaValue.Boolean(comparison >= 0);
        }
    }

    private static bool IsNumericKind(OdsFormulaValueKind kind) => kind == OdsFormulaValueKind.Number ||
        kind == OdsFormulaValueKind.Boolean || kind == OdsFormulaValueKind.Empty;

    private double RequireNumber(OdsFormulaValue value) {
        try { return value.AsNumber(); }
        catch (InvalidOperationException) { throw Error("A numeric value was required."); }
    }

    private OdsFormulaValue Number(double value) => double.IsNaN(value) || double.IsInfinity(value)
        ? OdsFormulaValue.Error("Formula produced a non-finite number.")
        : OdsFormulaValue.Number(value);

    private static OdsFormulaValue? Propagate(OdsFormulaValue left, OdsFormulaValue right) {
        if (left.Kind == OdsFormulaValueKind.Error) return left;
        if (right.Kind == OdsFormulaValueKind.Error) return right;
        return null;
    }

    private static List<OdsFormulaValue> Flatten(IEnumerable<OdsFormulaOperand> arguments) {
        var result = new List<OdsFormulaValue>();
        foreach (OdsFormulaOperand argument in arguments) result.AddRange(argument.Values);
        return result;
    }

    private OdsFormulaToken Take() {
        OdsFormulaToken current = _current;
        _current = _lexer.Next();
        return current;
    }

    private void Expect(OdsFormulaTokenKind kind) {
        if (_current.Kind != kind) throw Error("Expected '" + kind + "' but found '" + _current.Text + "'.");
        Take();
    }

    private OdsFormulaException Error(string message) => new OdsFormulaException(message);
    private void EnterSyntax() {
        _syntaxDepth++;
        if (_syntaxDepth > _context.Options.MaximumDependencyDepth) throw Error("Formula syntax depth limit exceeded.");
    }
    private static OdsFormulaOperand Scalar(OdsFormulaValue value) => OdsFormulaOperand.Scalar(value);
    private static bool IsComparison(OdsFormulaTokenKind kind) => kind >= OdsFormulaTokenKind.Equal && kind <= OdsFormulaTokenKind.GreaterOrEqual;
}

internal sealed class OdsFormulaOperand {
    private OdsFormulaOperand(IReadOnlyList<OdsFormulaValue> values) { Values = values; }
    internal IReadOnlyList<OdsFormulaValue> Values { get; }
    internal static OdsFormulaOperand Scalar(OdsFormulaValue value) => new OdsFormulaOperand(new[] { value });
    internal static OdsFormulaOperand Range(IReadOnlyList<OdsFormulaValue> values) => new OdsFormulaOperand(values);
    internal OdsFormulaValue RequireScalar() {
        if (Values.Count != 1) throw new OdsFormulaException("A range cannot be used as a scalar value in this OpenFormula subset.");
        return Values[0];
    }
}

internal readonly struct OdsFormulaReference {
    private OdsFormulaReference(string sheetName, long firstRow, long firstColumn, long lastRow, long lastColumn) {
        SheetName = sheetName; FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
    }
    internal string SheetName { get; }
    internal long FirstRow { get; }
    internal long FirstColumn { get; }
    internal long LastRow { get; }
    internal long LastColumn { get; }

    internal static OdsFormulaReference Parse(string lexical, string currentSheet) {
        string value = lexical.Trim();
        int colon = FindUnquoted(value, ':');
        string firstText = colon < 0 ? value : value.Substring(0, colon);
        string secondText = colon < 0 ? firstText : value.Substring(colon + 1);
        Endpoint first = ParseEndpoint(firstText, currentSheet);
        Endpoint second = ParseEndpoint(secondText, first.SheetName);
        if (!string.Equals(first.SheetName, second.SheetName, StringComparison.Ordinal)) {
            throw new OdsFormulaException("Three-dimensional formula ranges are not supported.");
        }
        return new OdsFormulaReference(first.SheetName, Math.Min(first.Row, second.Row), Math.Min(first.Column, second.Column),
            Math.Max(first.Row, second.Row), Math.Max(first.Column, second.Column));
    }

    private static Endpoint ParseEndpoint(string lexical, string defaultSheet) {
        string value = lexical.Trim();
        int dot = FindLastUnquoted(value, '.');
        string sheet = defaultSheet;
        string cell = value;
        if (dot >= 0) {
            string sheetPart = value.Substring(0, dot);
            cell = value.Substring(dot + 1);
            if (sheetPart.Length > 0) sheet = UnquoteSheet(sheetPart.TrimStart('$'));
        }
        cell = cell.Trim().TrimStart('$').Replace("$", string.Empty);
        int split = 0;
        while (split < cell.Length && char.IsLetter(cell[split])) split++;
        if (split == 0 || split == cell.Length || !long.TryParse(cell.Substring(split), NumberStyles.None, CultureInfo.InvariantCulture, out long row) || row < 1) {
            throw new OdsFormulaException("Invalid cell reference '" + lexical + "'.");
        }
        long column = 0;
        for (int index = 0; index < split; index++) {
            char character = char.ToUpperInvariant(cell[index]);
            if (character < 'A' || character > 'Z') throw new OdsFormulaException("Invalid cell reference '" + lexical + "'.");
            column = checked(column * 26L + character - 'A' + 1L);
        }
        return new Endpoint(sheet, row - 1L, column - 1L);
    }

    private static string UnquoteSheet(string value) {
        if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'') {
            return value.Substring(1, value.Length - 2).Replace("''", "'");
        }
        return value;
    }

    private static int FindUnquoted(string value, char sought) {
        bool quoted = false;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\'') {
                if (quoted && index + 1 < value.Length && value[index + 1] == '\'') { index++; continue; }
                quoted = !quoted;
            } else if (!quoted && value[index] == sought) return index;
        }
        return -1;
    }

    private static int FindLastUnquoted(string value, char sought) {
        bool quoted = false; int found = -1;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\'') {
                if (quoted && index + 1 < value.Length && value[index + 1] == '\'') { index++; continue; }
                quoted = !quoted;
            } else if (!quoted && value[index] == sought) found = index;
        }
        return found;
    }

    private readonly struct Endpoint {
        internal Endpoint(string sheetName, long row, long column) { SheetName = sheetName; Row = row; Column = column; }
        internal string SheetName { get; }
        internal long Row { get; }
        internal long Column { get; }
    }
}
