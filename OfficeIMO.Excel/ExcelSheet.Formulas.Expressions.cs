using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateScalarExpression(string expression, out FormulaArgumentValue value) {
            ExcelFormulaSyntaxCache? cache = ExcelFormulaSyntaxCache.Current;
            if (cache != null) {
                if (!cache.TryGetScalar(expression, out object? compiled)) {
                    compiled = ScalarExpressionCompiler.Compile(expression);
                    cache.AddScalar(expression, compiled);
                }
                if (compiled is ScalarExpressionNode root) return TryEvaluateCompiledScalarExpression(root, out value);
            }
            var parser = new ScalarFormulaParser(this, expression);
            return parser.TryParse(out value);
        }

        // Single-pass precedence parsing. References and functions use the existing evaluator;
        // syntax nesting and function recursion have separate bounded guards.
        private struct ScalarFormulaParser {
            private readonly ExcelSheet _sheet;
            private readonly string _text;
            private int _position;
            private int _depth;
            private bool _lastOperationWasAddition;
            private double _lastLeft;
            private double _lastRight;

            internal ScalarFormulaParser(ExcelSheet sheet, string text) {
                _sheet = sheet;
                _text = text;
                _position = 0;
                _depth = 0;
                _lastOperationWasAddition = false;
                _lastLeft = _lastRight = 0;
            }

            internal bool TryParse(out FormulaArgumentValue value) {
                if (!Parse(1, out value)) return false;
                SkipWhiteSpace();
                if (_position == _text.Length && _lastOperationWasAddition && value.Number.HasValue) {
                    value = CancelAdditionNoise(value, _lastLeft, _lastRight);
                }
                return _position == _text.Length && !value.IsUnresolvedFormula;
            }

            private bool Parse(int minimumPrecedence, out FormulaArgumentValue value) {
                value = default;
                if (!_sheet.HasSufficientFormulaExecutionStack()) return false;
                if (++_depth > 128) { _depth--; return false; }
                try {
                    SkipWhiteSpace();
                    if (_position == _text.Length) return false;
                    char first = _text[_position];
                    if (first == '+' || first == '-') {
                        _position++;
                        if (!Parse(7, out value)) return false;
                        value = NumericUnary(value, first == '-' ? -1 : 1);
                        _lastOperationWasAddition = false;
                    } else if (first == '(') {
                        _position++;
                        if (!Parse(1, out value)) return false;
                        SkipWhiteSpace();
                        if (_position == _text.Length || _text[_position++] != ')') return false;
                        _lastOperationWasAddition = false;
                    } else if (!ReadAtom(out value)) {
                        return false;
                    }

                    while (true) {
                        SkipWhiteSpace();
                        if (_position == _text.Length) break;
                        string op = ReadOperator();
                        int precedence = Precedence(op);
                        if (precedence < minimumPrecedence) break;
                        _position += op.Length;
                        if (op == "%") {
                            value = NumericUnary(value, 0.01);
                            _lastOperationWasAddition = false;
                            continue;
                        }
                        // Excel evaluates equal-precedence operators left to right, including ^.
                        if (!Parse(precedence + 1, out FormulaArgumentValue right)) return false;
                        if (!Apply(value, op, right, out value)) return false;
                    }
                    return !value.IsUnresolvedFormula;
                } finally { _depth--; }
            }

            private bool ReadAtom(out FormulaArgumentValue value) {
                value = default;
                if (!TryScanScalarAtom(_text, ref _position, out int start, out string? errorLiteral)) return false;
                if (errorLiteral != null) {
                    value = FormulaArgumentValue.Error(errorLiteral);
                    return true;
                }
                return _sheet.TryResolveFormulaArgument(_text.Substring(start, _position - start), out value,
                    allowScalarExpression: false);
            }

            internal static bool IsOperator(char ch) => ch == '+' || ch == '-' || ch == '*' || ch == '/'
                || ch == '^' || ch == '%' || ch == '&' || ch == '=' || ch == '<' || ch == '>';

            internal static readonly string[] ScalarErrorLiterals = { "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!" };

            private string ReadOperator() => ReadScalarOperator(_text, _position);

            internal static string ReadScalarOperator(string text, int position) {
                char ch = text[position];
                if ((ch == '<' || ch == '>') && position + 1 < text.Length) {
                    char next = text[position + 1];
                    if (next == '=') return ch == '<' ? "<=" : ">=";
                    if (ch == '<' && next == '>') return "<>";
                }
                switch (ch) {
                    case '+': return "+";
                    case '-': return "-";
                    case '*': return "*";
                    case '/': return "/";
                    case '^': return "^";
                    case '%': return "%";
                    case '&': return "&";
                    case '=': return "=";
                    case '<': return "<";
                    case '>': return ">";
                    default: return string.Empty;
                }
            }

            internal static int Precedence(string op) {
                switch (op) {
                    case "=": case "<>": case "<": case ">": case "<=": case ">=": return 1;
                    case "&": return 2;
                    case "+": case "-": return 3;
                    case "*": case "/": return 4;
                    case "^": return 5;
                    case "%": return 6;
                    default: return 0;
                }
            }

            private static bool TryNumber(FormulaArgumentValue value, out double number) {
                number = value.Number ?? 0;
                if (value.Number.HasValue || !value.HasValue) return true;
                return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
            }

            private static FormulaArgumentValue Numeric(double number) =>
                double.IsNaN(number) || double.IsInfinity(number) ? FormulaArgumentValue.Error("#NUM!")
                    : new FormulaArgumentValue(number, InvariantNumberText.Get(number));

            internal static FormulaArgumentValue NumericUnary(FormulaArgumentValue value, double factor) {
                if (value.IsError || value.IsUnresolvedFormula) return value;
                return TryNumber(value, out double number) ? Numeric(number * factor) : FormulaArgumentValue.Error("#VALUE!");
            }

            private bool Apply(FormulaArgumentValue left, string op, FormulaArgumentValue right,
                out FormulaArgumentValue value) =>
                ApplyOperator(left, op, right, out value, ref _lastOperationWasAddition, ref _lastLeft, ref _lastRight);

            internal static bool ApplyOperator(FormulaArgumentValue left, string op, FormulaArgumentValue right,
                out FormulaArgumentValue value, ref bool lastOperationWasAddition, ref double lastLeft, ref double lastRight) {
                value = default;
                lastOperationWasAddition = false;
                if (left.IsUnresolvedFormula || right.IsUnresolvedFormula) return false;
                if (left.IsError || right.IsError) { value = left.IsError ? left : right; return true; }
                if (op == "&") { value = new FormulaArgumentValue(null, ScalarText(left) + ScalarText(right)); return true; }
                if (Precedence(op) == 1) {
                    if (!TryCompareFormulaValues(left, op, right, out bool comparison)) return false;
                    value = new FormulaArgumentValue(comparison ? 1 : 0, comparison ? "1" : "0", isBoolean: true);
                    return true;
                }
                if (!TryNumber(left, out double a) || !TryNumber(right, out double b)) {
                    value = FormulaArgumentValue.Error("#VALUE!"); return true;
                }
                switch (op) {
                    case "+": value = Numeric(a + b); lastOperationWasAddition = true; break;
                    case "-": value = Numeric(a - b); lastOperationWasAddition = true; break;
                    case "*": value = Numeric(a * b); break;
                    case "/": value = b == 0 ? FormulaArgumentValue.Error("#DIV/0!") : Numeric(a / b); break;
                    case "^": value = Numeric(Math.Pow(a, b)); break;
                    default: return false;
                }
                lastLeft = a;
                lastRight = b;
                return true;
            }

            /// <summary>Folds a sum or difference that cancels to rounding noise to zero, as Excel displays it.</summary>
            internal static FormulaArgumentValue CancelAdditionNoise(FormulaArgumentValue value, double lastLeft, double lastRight) {
                double result = value.Number!.Value;
                double threshold = Math.Max(Math.Abs(lastLeft), Math.Abs(lastRight)) * (8d * 2.2204460492503131e-16);
                return result != 0d && Math.Abs(result) < threshold ? new FormulaArgumentValue(0d, "0") : value;
            }

            private void SkipWhiteSpace() {
                while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++;
            }

            private static string ScalarText(FormulaArgumentValue value) => FormulaValueToText(value);
        }
    }
}
