using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioShapeGeometry {

        private sealed class CellExpressionParser {
            private readonly string _text;
            private readonly double _width;
            private readonly double _height;
            private readonly double _locPinX;
            private readonly double _locPinY;
            private readonly double _pinX;
            private readonly double _pinY;
            private readonly double _angle;
            private int _index;

            private CellExpressionParser(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle) {
                _text = text;
                _width = width;
                _height = height;
                _locPinX = locPinX;
                _locPinY = locPinY;
                _pinX = pinX;
                _pinY = pinY;
                _angle = angle;
            }

            internal static bool TryEvaluate(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle, out double value) {
                CellExpressionParser parser = new(text, width, height, locPinX, locPinY, pinX, pinY, angle);
                if (parser.TryParseExpression(out value)) {
                    parser.SkipWhitespace();
                    if (parser._index == parser._text.Length) {
                        return true;
                    }
                }

                value = 0D;
                return false;
            }

            private bool TryParseExpression(out double value) {
                if (!TryParseTerm(out value)) {
                    return false;
                }

                while (true) {
                    SkipWhitespace();
                    if (TryRead('+')) {
                        if (!TryParseTerm(out double addend)) {
                            return false;
                        }

                        value += addend;
                    } else if (TryRead('-')) {
                        if (!TryParseTerm(out double subtrahend)) {
                            return false;
                        }

                        value -= subtrahend;
                    } else {
                        return true;
                    }
                }
            }

            private bool TryParseTerm(out double value) {
                if (!TryParsePower(out value)) {
                    return false;
                }

                while (true) {
                    SkipWhitespace();
                    if (TryRead('*')) {
                        if (!TryParsePower(out double factor)) {
                            return false;
                        }

                        value *= factor;
                    } else if (TryRead('/')) {
                        if (!TryParsePower(out double divisor) || Math.Abs(divisor) <= 1e-12) {
                            return false;
                        }

                        value /= divisor;
                    } else {
                        return true;
                    }
                }
            }

            private bool TryParsePower(out double value) {
                if (!TryParseFactor(out value)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead('^')) {
                    return true;
                }

                if (!TryParsePower(out double exponent)) {
                    return false;
                }

                value = ShapeSheetPower(value, exponent);
                return IsFinite(value);
            }

            private bool TryParseFactor(out double value) {
                SkipWhitespace();
                if (TryRead('+')) {
                    return TryParseFactor(out value);
                }

                if (TryRead('-')) {
                    if (!TryParseFactor(out value)) {
                        return false;
                    }

                    value = -value;
                    return true;
                }

                if (TryRead('(')) {
                    if (!TryParseExpression(out value)) {
                        return false;
                    }

                    SkipWhitespace();
                    return TryRead(')');
                }

                if (TryParseIdentifier(out value)) {
                    return true;
                }

                return TryParseNumber(out value);
            }

            private bool TryParseIdentifier(out double value) {
                SkipWhitespace();
                int start = _index;
                if (_index < _text.Length && char.IsLetter(_text[_index])) {
                    _index++;
                    while (_index < _text.Length && char.IsLetterOrDigit(_text[_index])) {
                        _index++;
                    }
                }

                if (_index == start) {
                    value = 0D;
                    return false;
                }

                string identifier = _text.Substring(start, _index - start);
                SkipWhitespace();
                if (_index < _text.Length && _text[_index] == '(') {
                    return TryParseFunction(identifier, out value);
                }

                if (string.Equals(identifier, "Width", StringComparison.OrdinalIgnoreCase)) {
                    value = _width;
                    return true;
                }

                if (string.Equals(identifier, "Height", StringComparison.OrdinalIgnoreCase)) {
                    value = _height;
                    return true;
                }

                if (string.Equals(identifier, "LocPinX", StringComparison.OrdinalIgnoreCase)) {
                    value = _locPinX;
                    return true;
                }

                if (string.Equals(identifier, "LocPinY", StringComparison.OrdinalIgnoreCase)) {
                    value = _locPinY;
                    return true;
                }

                if (string.Equals(identifier, "PinX", StringComparison.OrdinalIgnoreCase)) {
                    value = _pinX;
                    return true;
                }

                if (string.Equals(identifier, "PinY", StringComparison.OrdinalIgnoreCase)) {
                    value = _pinY;
                    return true;
                }

                if (string.Equals(identifier, "Angle", StringComparison.OrdinalIgnoreCase)) {
                    value = _angle;
                    return true;
                }

                if (string.Equals(identifier, "TRUE", StringComparison.OrdinalIgnoreCase)) {
                    value = 1D;
                    return true;
                }

                if (string.Equals(identifier, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                    value = 0D;
                    return true;
                }

                value = 0D;
                return false;
            }

            private bool TryParseFunction(string identifier, out double value) {
                value = 0D;
                if (string.Equals(identifier, "IF", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseIfFunction(out value);
                }

                if (string.Equals(identifier, "AND", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(identifier, "OR", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(identifier, "NOT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseLogicalFunction(identifier, out value);
                }

                if (string.Equals(identifier, "PI", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseNoArgumentFunction(Math.PI, out value);
                }

                if (string.Equals(identifier, "SIN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Sin, out value);
                }

                if (string.Equals(identifier, "COS", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Cos, out value);
                }

                if (string.Equals(identifier, "TAN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Tan, out value);
                }

                if (string.Equals(identifier, "ATAN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Atan, out value);
                }

                if (string.Equals(identifier, "ATAN2", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction((y, x) => Math.Abs(y) <= 1e-12 && Math.Abs(x) <= 1e-12 ? 0D : Math.Atan2(y, x), out value);
                }

                if (string.Equals(identifier, "RAD", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(angle => angle * Math.PI / 180D, out value);
                }

                if (string.Equals(identifier, "DEG", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(angle => angle * 180D / Math.PI, out value);
                }

                if (string.Equals(identifier, "ABS", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Abs, out value);
                }

                if (string.Equals(identifier, "SQRT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Sqrt, out value);
                }

                if (string.Equals(identifier, "INT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Floor, out value);
                }

                if (string.Equals(identifier, "POW", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction(ShapeSheetPower, out value);
                }

                if (string.Equals(identifier, "ROUND", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction(RoundShapeSheetValue, out value);
                }

                bool isMin = string.Equals(identifier, "MIN", StringComparison.OrdinalIgnoreCase);
                bool isMax = string.Equals(identifier, "MAX", StringComparison.OrdinalIgnoreCase);
                if (!isMin && !isMax) {
                    return false;
                }

                if (!TryRead('(')) {
                    return false;
                }

                bool hasValue = false;
                double result = isMin
                    ? double.PositiveInfinity
                    : double.NegativeInfinity;
                while (true) {
                    if (!TryParseExpression(out double argument)) {
                        return false;
                    }

                    hasValue = true;
                    result = isMin
                        ? Math.Min(result, argument)
                        : Math.Max(result, argument);
                    SkipWhitespace();
                    if (TryRead(')')) {
                        value = result;
                        return hasValue && IsFinite(value);
                    }

                    if (!TryRead(',')) {
                        return false;
                    }
                }
            }

            private bool TryParseNoArgumentFunction(double result, out double value) {
                value = 0D;
                if (!TryRead('(')) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = result;
                return IsFinite(value);
            }

            private bool TryParseSingleArgumentFunction(Func<double, double> function, out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseExpression(out double argument)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = function(argument);
                return IsFinite(value);
            }

            private bool TryParseTwoArgumentFunction(Func<double, double, double> function, out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseExpression(out double firstArgument) ||
                    !TryReadComma() ||
                    !TryParseExpression(out double secondArgument)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = function(firstArgument, secondArgument);
                return IsFinite(value);
            }

            private static double ShapeSheetPower(double number, double exponent) {
                if (Math.Abs(number) <= 1e-12 && exponent <= 0D) {
                    return 0D;
                }

                if (number < 0D && !IsNearlyInteger(exponent)) {
                    return 0D;
                }

                return Math.Pow(number, exponent);
            }

            private static double RoundShapeSheetValue(double number, double numberOfDigits) {
                int digits = (int)Math.Round(numberOfDigits, MidpointRounding.AwayFromZero);
                double factor = Math.Pow(10D, Math.Abs(digits));
                if (!IsFinite(factor) || Math.Abs(factor) <= 1e-12) {
                    return double.NaN;
                }

                if (digits >= 0) {
                    return Math.Round(number * factor, 0, MidpointRounding.AwayFromZero) / factor;
                }

                return Math.Round(number / factor, 0, MidpointRounding.AwayFromZero) * factor;
            }

            private static bool IsNearlyInteger(double value) =>
                Math.Abs(value - Math.Round(value)) <= 1e-9;

            private bool TryParseLogicalFunction(string identifier, out double value) {
                value = 0D;
                if (!TryReadFunctionArguments(out List<string> arguments)) {
                    return false;
                }

                bool result;
                if (string.Equals(identifier, "AND", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count == 0) {
                        return false;
                    }

                    result = true;
                    foreach (string argument in arguments) {
                        if (!TryEvaluateCondition(argument, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                            return false;
                        }

                        result &= argumentValue;
                    }
                } else if (string.Equals(identifier, "OR", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count == 0) {
                        return false;
                    }

                    result = false;
                    foreach (string argument in arguments) {
                        if (!TryEvaluateCondition(argument, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                            return false;
                        }

                        result |= argumentValue;
                    }
                } else if (string.Equals(identifier, "NOT", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count != 1 ||
                        !TryEvaluateCondition(arguments[0], _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                        return false;
                    }

                    result = !argumentValue;
                } else {
                    return false;
                }

                value = result ? 1D : 0D;
                return true;
            }

            private bool TryParseIfFunction(out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseCondition(out bool condition) ||
                    !TryReadComma() ||
                    !TryReadFunctionArgument(',', out string? whenTrueExpression) ||
                    !TryReadFunctionArgument(')', out string? whenFalseExpression)) {
                    return false;
                }

                string selectedExpression = condition ? whenTrueExpression! : whenFalseExpression!;
                return TryEvaluate(selectedExpression, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out value) &&
                       IsFinite(value);
            }

            private static bool TryEvaluateCondition(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle, out bool value) {
                CellExpressionParser parser = new(text, width, height, locPinX, locPinY, pinX, pinY, angle);
                if (parser.TryParseCondition(out value)) {
                    parser.SkipWhitespace();
                    if (parser._index == parser._text.Length) {
                        return true;
                    }
                }

                value = false;
                return false;
            }

            private bool TryParseCondition(out bool value) {
                value = false;
                if (!TryParseExpression(out double left)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryReadComparisonOperator(out string? comparisonOperator)) {
                    value = Math.Abs(left) > 1e-9;
                    return true;
                }

                if (!TryParseExpression(out double right)) {
                    return false;
                }

                switch (comparisonOperator) {
                    case "<":
                        value = left < right;
                        return true;
                    case "<=":
                        value = left <= right;
                        return true;
                    case ">":
                        value = left > right;
                        return true;
                    case ">=":
                        value = left >= right;
                        return true;
                    case "=":
                        value = Math.Abs(left - right) <= 1e-9;
                        return true;
                    case "<>":
                    case "!=":
                        value = Math.Abs(left - right) > 1e-9;
                        return true;
                    default:
                        return false;
                }
            }

            private bool TryReadComparisonOperator(out string? comparisonOperator) {
                comparisonOperator = null;
                if (_index >= _text.Length) {
                    return false;
                }

                if (_index + 1 < _text.Length) {
                    string pair = _text.Substring(_index, 2);
                    if (pair == "<=" || pair == ">=" || pair == "<>" || pair == "!=") {
                        comparisonOperator = pair;
                        _index += 2;
                        return true;
                    }
                }

                char current = _text[_index];
                if (current == '<' || current == '>' || current == '=') {
                    comparisonOperator = current.ToString();
                    _index++;
                    return true;
                }

                return false;
            }

            private bool TryReadComma() {
                SkipWhitespace();
                return TryRead(',');
            }

            private bool TryReadFunctionArguments(out List<string> arguments) {
                arguments = new List<string>();
                if (!TryRead('(')) {
                    return false;
                }

                int start = _index;
                int depth = 0;
                while (_index < _text.Length) {
                    char current = _text[_index];
                    if (current == '(') {
                        depth++;
                    } else if (current == ')') {
                        if (depth == 0) {
                            string argument = _text.Substring(start, _index - start).Trim();
                            if (argument.Length == 0) {
                                return false;
                            }

                            arguments.Add(argument);
                            _index++;
                            return true;
                        }

                        depth--;
                    } else if (current == ',' && depth == 0) {
                        string argument = _text.Substring(start, _index - start).Trim();
                        if (argument.Length == 0) {
                            return false;
                        }

                        arguments.Add(argument);
                        _index++;
                        start = _index;
                        continue;
                    }

                    _index++;
                }

                return false;
            }

            private bool TryReadFunctionArgument(char delimiter, out string? argument) {
                SkipWhitespace();
                int start = _index;
                int depth = 0;
                while (_index < _text.Length) {
                    char current = _text[_index];
                    if (current == '(') {
                        depth++;
                    } else if (current == ')') {
                        if (depth == 0) {
                            if (delimiter != ')') {
                                argument = null;
                                return false;
                            }

                            argument = _text.Substring(start, _index - start).Trim();
                            _index++;
                            return argument.Length > 0;
                        }

                        depth--;
                    } else if (current == ',' && depth == 0) {
                        if (delimiter != ',') {
                            argument = null;
                            return false;
                        }

                        argument = _text.Substring(start, _index - start).Trim();
                        _index++;
                        return argument.Length > 0;
                    }

                    _index++;
                }

                argument = null;
                return false;
            }

            private bool TryParseNumber(out double value) {
                SkipWhitespace();
                int start = _index;
                bool hasDigit = false;
                while (_index < _text.Length && char.IsDigit(_text[_index])) {
                    _index++;
                    hasDigit = true;
                }

                if (_index < _text.Length && _text[_index] == '.') {
                    _index++;
                    while (_index < _text.Length && char.IsDigit(_text[_index])) {
                        _index++;
                        hasDigit = true;
                    }
                }

                if (!hasDigit) {
                    value = 0D;
                    return false;
                }

                if (_index < _text.Length && (_text[_index] == 'e' || _text[_index] == 'E')) {
                    int exponentStart = _index;
                    _index++;
                    if (_index < _text.Length && (_text[_index] == '+' || _text[_index] == '-')) {
                        _index++;
                    }

                    int exponentDigitsStart = _index;
                    while (_index < _text.Length && char.IsDigit(_text[_index])) {
                        _index++;
                    }

                    if (_index == exponentDigitsStart) {
                        _index = exponentStart;
                    }
                }

                if (!double.TryParse(_text.Substring(start, _index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                    return false;
                }

                TryApplyUnitSuffix(ref value);
                return IsFinite(value);
            }

            private void TryApplyUnitSuffix(ref double value) {
                int suffixStart = _index;
                SkipWhitespace();
                if (_index < _text.Length && _text[_index] == '%') {
                    value *= 0.01D;
                    _index++;
                    return;
                }

                int unitStart = _index;
                while (_index < _text.Length && char.IsLetter(_text[_index])) {
                    _index++;
                }

                if (_index == unitStart) {
                    _index = suffixStart;
                    return;
                }

                string unit = _text.Substring(unitStart, _index - unitStart);
                if (TryGetUnitScale(unit, out double scale)) {
                    value *= scale;
                    return;
                }

                _index = suffixStart;
            }

            private static bool TryGetUnitScale(string unit, out double scale) {
                switch (unit.ToLowerInvariant()) {
                    case "deg":
                    case "degree":
                    case "degrees":
                        scale = Math.PI / 180D;
                        return true;
                    case "rad":
                    case "radian":
                    case "radians":
                    case "in":
                    case "inch":
                    case "inches":
                        scale = 1D;
                        return true;
                    case "ft":
                    case "foot":
                    case "feet":
                        scale = 12D;
                        return true;
                    case "yd":
                    case "yard":
                    case "yards":
                        scale = 36D;
                        return true;
                    case "mi":
                    case "mile":
                    case "miles":
                        scale = 63360D;
                        return true;
                    case "mm":
                        scale = 1D / 25.4D;
                        return true;
                    case "cm":
                        scale = 1D / 2.54D;
                        return true;
                    case "m":
                    case "meter":
                    case "meters":
                        scale = 100D / 2.54D;
                        return true;
                    case "km":
                    case "kilometer":
                    case "kilometers":
                        scale = 100000D / 2.54D;
                        return true;
                    case "pt":
                    case "point":
                    case "points":
                        scale = 1D / 72D;
                        return true;
                    case "pc":
                    case "pica":
                    case "picas":
                        scale = 1D / 6D;
                        return true;
                    default:
                        scale = 1D;
                        return false;
                }
            }

            private bool TryRead(char expected) {
                if (_index < _text.Length && _text[_index] == expected) {
                    _index++;
                    return true;
                }

                return false;
            }

            private void SkipWhitespace() {
                while (_index < _text.Length && char.IsWhiteSpace(_text[_index])) {
                    _index++;
                }
            }
        }
    }
}
