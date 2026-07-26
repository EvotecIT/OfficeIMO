using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Evaluates the bounded CSS Values length-math subset shared by layout and paint parsers.
/// The parser keeps number and length dimensions distinct so malformed expressions fail
/// instead of silently producing pixels.
/// </summary>
internal sealed class HtmlCssLengthMathEvaluator {
    private const int MaximumExpressionLength = 4096;
    private const int MaximumNestingDepth = 32;
    private const int MaximumOperations = 256;

    private readonly string _text;
    private readonly double _reference;
    private readonly double _fontSize;
    private readonly double _rootFontSize;
    private int _index;
    private int _depth;
    private int _operations;

    private HtmlCssLengthMathEvaluator(
        string text,
        double reference,
        double fontSize,
        double rootFontSize) {
        _text = text;
        _reference = reference;
        _fontSize = fontSize;
        _rootFontSize = rootFontSize;
    }

    internal static bool TryEvaluate(
        string value,
        double reference,
        double fontSize,
        double rootFontSize,
        out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaximumExpressionLength) return false;

        var parser = new HtmlCssLengthMathEvaluator(value, reference, fontSize, rootFontSize);
        if (!parser.TryParseExpression(out CssNumeric resolved)) return false;
        parser.SkipWhitespace();
        if (parser._index != parser._text.Length
            || resolved.Dimension == CssNumericDimension.Number && resolved.Value != 0D
            || !IsFinite(resolved.Value)) {
            return false;
        }

        result = resolved.Value;
        return true;
    }

    private bool TryParseExpression(out CssNumeric result) {
        if (!TryParseProduct(out result)) return false;
        while (true) {
            SkipWhitespace();
            if (!TryReadOperator('+') && !TryReadOperator('-')) return true;
            char operation = _text[_index - 1];
            if (!RecordOperation() || !TryParseProduct(out CssNumeric right)) return false;
            if (!TryAddOrSubtract(result, right, operation, out result)) return false;
        }
    }

    private bool TryParseProduct(out CssNumeric result) {
        if (!TryParseUnary(out result)) return false;
        while (true) {
            SkipWhitespace();
            if (!TryReadOperator('*') && !TryReadOperator('/')) return true;
            char operation = _text[_index - 1];
            if (!RecordOperation() || !TryParseUnary(out CssNumeric right)) return false;
            if (!TryMultiplyOrDivide(result, right, operation, out result)) return false;
        }
    }

    private bool TryParseUnary(out CssNumeric result) {
        SkipWhitespace();
        bool negative = false;
        if (TryReadOperator('+')) {
            SkipWhitespace();
        } else if (TryReadOperator('-')) {
            negative = true;
            SkipWhitespace();
        }

        if (!TryParsePrimary(out result)) return false;
        if (negative) result = new CssNumeric(-result.Value, result.Dimension);
        return IsFinite(result.Value);
    }

    private bool TryParsePrimary(out CssNumeric result) {
        result = default;
        SkipWhitespace();
        if (_index >= _text.Length) return false;

        if (_text[_index] == '(') {
            _index++;
            if (!EnterNesting() || !TryParseExpression(out result)) return false;
            SkipWhitespace();
            if (!TryReadOperator(')')) return false;
            ExitNesting();
            return true;
        }

        int identifierStart = _index;
        while (_index < _text.Length && (char.IsLetter(_text[_index]) || _text[_index] == '-')) _index++;
        if (_index > identifierStart) {
            string function = _text.Substring(identifierStart, _index - identifierStart).ToLowerInvariant();
            SkipWhitespace();
            if (!TryReadOperator('(')) return false;
            if (!EnterNesting()) return false;
            bool success = TryParseFunction(function, out result);
            ExitNesting();
            return success;
        }

        return TryParseNumeric(out result);
    }

    private bool TryParseFunction(string function, out CssNumeric result) {
        result = default;
        if (function == "calc") {
            if (!TryParseExpression(out result)) return false;
            SkipWhitespace();
            return TryReadOperator(')');
        }

        if (function != "min" && function != "max" && function != "clamp") return false;

        var arguments = new List<CssNumeric>();
        while (true) {
            if (!TryParseExpression(out CssNumeric argument)) return false;
            arguments.Add(argument);
            SkipWhitespace();
            if (TryReadOperator(')')) break;
            if (!TryReadOperator(',')) return false;
        }

        if (function == "clamp") {
            if (arguments.Count != 3
                || !TryNormalizeDimensions(arguments, out CssNumericDimension dimension)) return false;
            double minimum = arguments[0].Value;
            double preferred = arguments[1].Value;
            double maximum = arguments[2].Value;
            result = new CssNumeric(Math.Max(minimum, Math.Min(preferred, maximum)), dimension);
            return IsFinite(result.Value);
        }

        if (arguments.Count == 0 || !TryNormalizeDimensions(arguments, out CssNumericDimension aggregateDimension)) return false;
        double value = arguments[0].Value;
        for (int index = 1; index < arguments.Count; index++) {
            value = function == "min"
                ? Math.Min(value, arguments[index].Value)
                : Math.Max(value, arguments[index].Value);
        }
        result = new CssNumeric(value, aggregateDimension);
        return IsFinite(result.Value);
    }

    private bool TryParseNumeric(out CssNumeric result) {
        result = default;
        int start = _index;
        bool hasDigit = false;
        while (_index < _text.Length && char.IsDigit(_text[_index])) {
            hasDigit = true;
            _index++;
        }
        if (_index < _text.Length && _text[_index] == '.') {
            _index++;
            while (_index < _text.Length && char.IsDigit(_text[_index])) {
                hasDigit = true;
                _index++;
            }
        }
        if (!hasDigit) return false;

        if (_index < _text.Length && (_text[_index] == 'e' || _text[_index] == 'E')) {
            int exponentStart = _index++;
            if (_index < _text.Length && (_text[_index] == '+' || _text[_index] == '-')) _index++;
            int exponentDigits = _index;
            while (_index < _text.Length && char.IsDigit(_text[_index])) _index++;
            if (exponentDigits == _index) {
                _index = exponentStart;
            }
        }

        if (!double.TryParse(
                _text.Substring(start, _index - start),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double number)
            || !IsFinite(number)) {
            return false;
        }

        int unitStart = _index;
        while (_index < _text.Length && (char.IsLetter(_text[_index]) || _text[_index] == '%')) _index++;
        string unit = _text.Substring(unitStart, _index - unitStart).ToLowerInvariant();
        if (unit.Length == 0) {
            result = new CssNumeric(number, CssNumericDimension.Number);
            return true;
        }

        double multiplier;
        switch (unit) {
            case "px": multiplier = 1D; break;
            case "pt": multiplier = HtmlRenderOptions.CssPixelsPerInch / 72D; break;
            case "pc": multiplier = HtmlRenderOptions.CssPixelsPerInch / 6D; break;
            case "in": multiplier = HtmlRenderOptions.CssPixelsPerInch; break;
            case "cm": multiplier = HtmlRenderOptions.CssPixelsPerInch / 2.54D; break;
            case "mm": multiplier = HtmlRenderOptions.CssPixelsPerInch / 25.4D; break;
            case "q": multiplier = HtmlRenderOptions.CssPixelsPerInch / 101.6D; break;
            case "em": multiplier = _fontSize; break;
            case "rem": multiplier = _rootFontSize; break;
            case "%": multiplier = _reference / 100D; break;
            default: return false;
        }

        result = new CssNumeric(number * multiplier, CssNumericDimension.Length);
        return IsFinite(result.Value);
    }

    private bool EnterNesting() {
        _depth++;
        return _depth <= MaximumNestingDepth;
    }

    private void ExitNesting() {
        if (_depth > 0) _depth--;
    }

    private bool RecordOperation() {
        _operations++;
        return _operations <= MaximumOperations;
    }

    private bool TryReadOperator(char value) {
        if (_index >= _text.Length || _text[_index] != value) return false;
        _index++;
        return true;
    }

    private void SkipWhitespace() {
        while (_index < _text.Length && char.IsWhiteSpace(_text[_index])) _index++;
    }

    private static bool TryAddOrSubtract(
        CssNumeric left,
        CssNumeric right,
        char operation,
        out CssNumeric result) {
        result = default;
        if (!TryUnifyDimensions(left, right, out CssNumericDimension dimension)) return false;
        double value = operation == '+' ? left.Value + right.Value : left.Value - right.Value;
        if (!IsFinite(value)) return false;
        result = new CssNumeric(value, dimension);
        return true;
    }

    private static bool TryMultiplyOrDivide(
        CssNumeric left,
        CssNumeric right,
        char operation,
        out CssNumeric result) {
        result = default;
        if (operation == '/') {
            if (right.Value == 0D) return false;
            CssNumericDimension quotientDimension;
            if (right.Dimension == CssNumericDimension.Number) {
                quotientDimension = left.Dimension;
            } else if (left.Dimension == CssNumericDimension.Length
                && right.Dimension == CssNumericDimension.Length) {
                quotientDimension = CssNumericDimension.Number;
            } else {
                return false;
            }
            double quotient = left.Value / right.Value;
            if (!IsFinite(quotient)) return false;
            result = new CssNumeric(quotient, quotientDimension);
            return true;
        }

        if (left.Dimension == CssNumericDimension.Length && right.Dimension == CssNumericDimension.Length) return false;
        CssNumericDimension dimension = left.Dimension == CssNumericDimension.Length || right.Dimension == CssNumericDimension.Length
            ? CssNumericDimension.Length
            : CssNumericDimension.Number;
        double product = left.Value * right.Value;
        if (!IsFinite(product)) return false;
        result = new CssNumeric(product, dimension);
        return true;
    }

    private static bool TryNormalizeDimensions(
        IReadOnlyList<CssNumeric> values,
        out CssNumericDimension dimension) {
        dimension = values.Any(value => value.Dimension == CssNumericDimension.Length)
            ? CssNumericDimension.Length
            : CssNumericDimension.Number;
        foreach (CssNumeric value in values) {
            if (value.Dimension != dimension && value.Value != 0D) return false;
        }
        return true;
    }

    private static bool TryUnifyDimensions(
        CssNumeric left,
        CssNumeric right,
        out CssNumericDimension dimension) {
        if (left.Dimension == right.Dimension) {
            dimension = left.Dimension;
            return true;
        }
        if (left.Dimension == CssNumericDimension.Number && left.Value == 0D) {
            dimension = right.Dimension;
            return true;
        }
        if (right.Dimension == CssNumericDimension.Number && right.Value == 0D) {
            dimension = left.Dimension;
            return true;
        }
        dimension = default;
        return false;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private readonly struct CssNumeric {
        internal CssNumeric(double value, CssNumericDimension dimension) {
            Value = value;
            Dimension = dimension;
        }

        internal double Value { get; }
        internal CssNumericDimension Dimension { get; }
    }

    private enum CssNumericDimension {
        Number,
        Length
    }
}
