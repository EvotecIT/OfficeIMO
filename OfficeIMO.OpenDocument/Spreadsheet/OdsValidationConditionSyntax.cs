using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace OfficeIMO.OpenDocument;

/// <summary>Value category constrained by an OpenDocument spreadsheet validation condition.</summary>
public enum OdsValidationValueKind {
    /// <summary>Whole-number cell content.</summary>
    WholeNumber,
    /// <summary>Decimal-number cell content.</summary>
    DecimalNumber,
    /// <summary>Date cell content.</summary>
    Date,
    /// <summary>Time cell content.</summary>
    Time,
    /// <summary>Length of text cell content.</summary>
    TextLength,
    /// <summary>An explicit list of text values.</summary>
    List
}

/// <summary>Comparison performed by an OpenDocument validation condition.</summary>
public enum OdsValidationComparison {
    /// <summary>Equal.</summary>
    Equal,
    /// <summary>Not equal.</summary>
    NotEqual,
    /// <summary>Less than.</summary>
    LessThan,
    /// <summary>Less than or equal.</summary>
    LessThanOrEqual,
    /// <summary>Greater than.</summary>
    GreaterThan,
    /// <summary>Greater than or equal.</summary>
    GreaterThanOrEqual,
    /// <summary>Between two values, inclusive.</summary>
    Between,
    /// <summary>Outside two values.</summary>
    NotBetween
}

/// <summary>
/// Typed syntax for the interoperable ODF validation-condition grammar. Unsupported implementation-specific
/// expressions fail closed instead of being guessed from substrings.
/// </summary>
public sealed class OdsValidationConditionSyntax {
    private readonly IReadOnlyList<string> _listValues;

    private OdsValidationConditionSyntax(
        OdsValidationValueKind valueKind,
        OdsValidationComparison? comparison,
        string? firstOperand,
        string? secondOperand,
        IReadOnlyList<string>? listValues) {
        ValueKind = valueKind;
        Comparison = comparison;
        FirstOperand = firstOperand;
        SecondOperand = secondOperand;
        _listValues = listValues ?? Array.Empty<string>();
    }

    /// <summary>Constrained value category.</summary>
    public OdsValidationValueKind ValueKind { get; }
    /// <summary>Comparison operation, or null for a list condition.</summary>
    public OdsValidationComparison? Comparison { get; }
    /// <summary>First authored scalar/formula operand.</summary>
    public string? FirstOperand { get; }
    /// <summary>Second authored scalar/formula operand for between conditions.</summary>
    public string? SecondOperand { get; }
    /// <summary>Decoded list values for a list condition.</summary>
    public IReadOnlyList<string> ListValues => _listValues;

    /// <summary>Creates a scalar validation condition.</summary>
    public static OdsValidationConditionSyntax Create(
        OdsValidationValueKind valueKind,
        OdsValidationComparison comparison,
        string firstOperand,
        string? secondOperand = null) {
        if (valueKind == OdsValidationValueKind.List) throw new ArgumentException("Use CreateList for list validation.", nameof(valueKind));
        if (string.IsNullOrWhiteSpace(firstOperand)) throw new ArgumentException("A first operand is required.", nameof(firstOperand));
        bool requiresSecond = comparison == OdsValidationComparison.Between || comparison == OdsValidationComparison.NotBetween;
        if (requiresSecond != !string.IsNullOrWhiteSpace(secondOperand)) {
            throw new ArgumentException(requiresSecond ? "This comparison requires two operands." : "This comparison accepts one operand.", nameof(secondOperand));
        }
        return new OdsValidationConditionSyntax(valueKind, comparison, firstOperand.Trim(), secondOperand?.Trim(), null);
    }

    /// <summary>Creates an explicit-list validation condition.</summary>
    public static OdsValidationConditionSyntax CreateList(IEnumerable<string> values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        var list = new List<string>();
        foreach (string? value in values) list.Add(value ?? string.Empty);
        if (list.Count == 0) throw new ArgumentException("At least one list value is required.", nameof(values));
        return new OdsValidationConditionSyntax(
            OdsValidationValueKind.List, null, null, null, new ReadOnlyCollection<string>(list));
    }

    /// <summary>Parses a complete interoperable validation condition.</summary>
    public static OdsValidationConditionSyntax Parse(string text) {
        if (!TryParse(text, out OdsValidationConditionSyntax? condition)) {
            throw new FormatException($"'{text}' is not a supported OpenDocument validation condition.");
        }
        return condition!;
    }

    /// <summary>Attempts to parse a complete interoperable validation condition.</summary>
    public static bool TryParse(string? text, out OdsValidationConditionSyntax? condition) {
        condition = null;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string value = text!.Trim();
        if (value.StartsWith("of:", StringComparison.OrdinalIgnoreCase)) value = value.Substring(3).Trim();

        const string listPrefix = "cell-content-is-in-list(";
        if (StartsWith(value, listPrefix) && value.EndsWith(")", StringComparison.Ordinal)) {
            string body = value.Substring(listPrefix.Length, value.Length - listPrefix.Length - 1);
            if (!TryParseStringList(body, out IReadOnlyList<string>? values)) return false;
            condition = CreateList(values!);
            return true;
        }

        if (TryParseTextLength(value, out condition)) return true;

        foreach ((string Prefix, OdsValidationValueKind Kind) item in TypePrefixes) {
            if (!StartsWith(value, item.Prefix)) continue;
            string remainder = value.Substring(item.Prefix.Length).Trim();
            if (!remainder.StartsWith("and", StringComparison.OrdinalIgnoreCase)
                || (remainder.Length > 3 && !char.IsWhiteSpace(remainder[3]))) return false;
            return TryParseScalar(item.Kind, remainder.Substring(3).Trim(), out condition);
        }
        return false;
    }

    /// <inheritdoc />
    public override string ToString() {
        if (ValueKind == OdsValidationValueKind.List) {
            return "of:cell-content-is-in-list(" + string.Join(";", ListValues.Select(Quote)) + ")";
        }

        string subject = ValueKind == OdsValidationValueKind.TextLength
            ? "cell-content-text-length"
            : "cell-content";
        string predicate;
        if (Comparison == OdsValidationComparison.Between || Comparison == OdsValidationComparison.NotBetween) {
            predicate = subject + (Comparison == OdsValidationComparison.Between ? "-is-between(" : "-is-not-between(")
                + FirstOperand + "," + SecondOperand + ")";
        } else {
            // LibreOffice's condition lexer accepts whitespace after an operator but not before it.
            // Emit the compact canonical spelling used by its own ODF exporter.
            predicate = subject + "()" + FormatOperator(Comparison!.Value) + FirstOperand;
        }
        if (ValueKind == OdsValidationValueKind.TextLength) return "of:" + predicate;
        return "of:" + FormatTypePrefix(ValueKind) + " and " + predicate;
    }

    private static readonly (string Prefix, OdsValidationValueKind Kind)[] TypePrefixes = {
        ("cell-content-is-whole-number()", OdsValidationValueKind.WholeNumber),
        ("cell-content-is-decimal-number()", OdsValidationValueKind.DecimalNumber),
        ("cell-content-is-date()", OdsValidationValueKind.Date),
        ("cell-content-is-time()", OdsValidationValueKind.Time)
    };

    private static bool TryParseTextLength(string value, out OdsValidationConditionSyntax? condition) {
        condition = null;
        if (!value.StartsWith("cell-content-text-length", StringComparison.OrdinalIgnoreCase)) return false;
        return TryParseScalar(OdsValidationValueKind.TextLength, value, out condition);
    }

    private static bool TryParseScalar(
        OdsValidationValueKind kind,
        string predicate,
        out OdsValidationConditionSyntax? condition) {
        condition = null;
        string subject = kind == OdsValidationValueKind.TextLength ? "cell-content-text-length" : "cell-content";
        string between = subject + "-is-between(";
        string notBetween = subject + "-is-not-between(";
        if ((StartsWith(predicate, between) || StartsWith(predicate, notBetween))
            && predicate.EndsWith(")", StringComparison.Ordinal)) {
            bool negate = StartsWith(predicate, notBetween);
            string prefix = negate ? notBetween : between;
            string body = predicate.Substring(prefix.Length, predicate.Length - prefix.Length - 1);
            if (!TrySplitOperands(body, out string? first, out string? second)) return false;
            condition = Create(kind, negate ? OdsValidationComparison.NotBetween : OdsValidationComparison.Between, first!, second);
            return true;
        }

        string comparisonPrefix = subject + "()";
        if (!StartsWith(predicate, comparisonPrefix)) return false;
        string remainder = predicate.Substring(comparisonPrefix.Length).Trim();
        foreach ((string Token, OdsValidationComparison Comparison) item in Operators) {
            if (!remainder.StartsWith(item.Token, StringComparison.Ordinal)) continue;
            string operand = remainder.Substring(item.Token.Length).Trim();
            if (operand.Length == 0) return false;
            condition = Create(kind, item.Comparison, operand);
            return true;
        }
        return false;
    }

    private static readonly (string Token, OdsValidationComparison Comparison)[] Operators = {
        ("!=", OdsValidationComparison.NotEqual), ("<=", OdsValidationComparison.LessThanOrEqual),
        (">=", OdsValidationComparison.GreaterThanOrEqual), ("=", OdsValidationComparison.Equal),
        ("<", OdsValidationComparison.LessThan), (">", OdsValidationComparison.GreaterThan)
    };

    private static bool TrySplitOperands(string value, out string? first, out string? second) {
        first = null;
        second = null;
        int depth = 0;
        bool quoted = false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quoted) {
                if (current == '"') {
                    if (index + 1 < value.Length && value[index + 1] == '"') { index++; continue; }
                    quoted = false;
                }
                continue;
            }
            if (current == '"') { quoted = true; continue; }
            if (current == '(' || current == '[') { depth++; continue; }
            if (current == ')' || current == ']') { if (depth == 0) return false; depth--; continue; }
            if (depth == 0 && current == ',') {
                if (first != null) return false;
                first = value.Substring(0, index).Trim();
                second = value.Substring(index + 1).Trim();
                break;
            }
        }
        return !quoted && depth == 0 && !string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second);
    }

    private static bool TryParseStringList(string value, out IReadOnlyList<string>? values) {
        var result = new List<string>();
        int cursor = 0;
        while (cursor < value.Length) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
            if (cursor >= value.Length || value[cursor] != '"') { values = null; return false; }
            cursor++;
            var item = new StringBuilder();
            bool closed = false;
            while (cursor < value.Length) {
                if (value[cursor] != '"') { item.Append(value[cursor++]); continue; }
                if (cursor + 1 < value.Length && value[cursor + 1] == '"') { item.Append('"'); cursor += 2; continue; }
                cursor++;
                closed = true;
                break;
            }
            if (!closed) { values = null; return false; }
            result.Add(item.ToString());
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
            if (cursor == value.Length) break;
            if (value[cursor] != ';') { values = null; return false; }
            cursor++;
        }
        values = result.Count == 0 ? null : new ReadOnlyCollection<string>(result);
        return values != null;
    }

    private static bool StartsWith(string value, string prefix) => value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
    private static string Quote(string value) => "\"" + value.Replace("\"", "\"\"") + "\"";

    private static string FormatTypePrefix(OdsValidationValueKind kind) => kind switch {
        OdsValidationValueKind.WholeNumber => "cell-content-is-whole-number()",
        OdsValidationValueKind.DecimalNumber => "cell-content-is-decimal-number()",
        OdsValidationValueKind.Date => "cell-content-is-date()",
        OdsValidationValueKind.Time => "cell-content-is-time()",
        _ => throw new InvalidOperationException("The validation kind does not use a type prefix.")
    };

    private static string FormatOperator(OdsValidationComparison comparison) => comparison switch {
        OdsValidationComparison.Equal => "=",
        OdsValidationComparison.NotEqual => "!=",
        OdsValidationComparison.LessThan => "<",
        OdsValidationComparison.LessThanOrEqual => "<=",
        OdsValidationComparison.GreaterThan => ">",
        OdsValidationComparison.GreaterThanOrEqual => ">=",
        _ => throw new InvalidOperationException("Between comparisons use function syntax.")
    };
}
