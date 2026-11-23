using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Scans one scalar operand from <paramref name="position"/>: an error literal, or text up to the next
        /// top-level operator or closing parenthesis. Shared by the interpreting parser and the compiler, so both
        /// split an expression at the same places.
        /// </summary>
        private static bool TryScanScalarAtom(string text, ref int position, out int start, out string? errorLiteral) {
            errorLiteral = null;
            start = position;
            if (text[start] == '#') {
                foreach (string error in ScalarFormulaParser.ScalarErrorLiterals) {
                    if (text.Length - start >= error.Length && string.Compare(text, start, error, 0, error.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                        position += error.Length;
                        errorLiteral = error;
                        return true;
                    }
                }
            }
            int parentheses = 0, brackets = 0;
            char quote = '\0';
            while (position < text.Length) {
                char ch = text[position];
                if (quote == '\0' && brackets > 0 && ch == '\'' && position + 1 < text.Length) {
                    position += 2;
                    continue;
                }
                if (quote != '\0') {
                    if (ch == quote) {
                        if (position + 1 < text.Length && text[position + 1] == quote) {
                            position += 2;
                            continue;
                        }
                        quote = '\0';
                    }
                } else if (brackets == 0 && (ch == '"' || ch == '\'')) {
                    quote = ch;
                } else if (ch == '[') { brackets++; }
                else if (ch == ']') { if (--brackets < 0) return false; }
                else if (brackets == 0) {
                    if (ch == '(') { if (++parentheses > 128) return false; }
                    else if (ch == ')') { if (parentheses == 0) break; parentheses--; }
                    else if (parentheses == 0 && ScalarFormulaParser.IsOperator(ch)) {
                        // A sign in a scientific numeric literal belongs to the literal.
                        bool exponentSign = (ch == '+' || ch == '-') && position > start
                            && (text[position - 1] == 'e' || text[position - 1] == 'E')
                            && OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text.Substring(start, position - start - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double _);
                        if (!exponentSign) break;
                    }
                }
                position++;
            }
            return quote == '\0' && parentheses == 0 && brackets == 0 && start != position;
        }

        private enum ScalarExpressionNodeKind {
            Level,
            Atom,
            Unary,
            Parentheses,
            Percent,
            Binary
        }

        /// <summary>
        /// One node of a compiled scalar expression. A <see cref="ScalarExpressionNodeKind.Level"/> node marks where
        /// the interpreting parser starts a nested precedence level, which checks the stack and rejects an
        /// unresolved result exactly where the parser would.
        /// </summary>
        private sealed class ScalarExpressionNode {
            private ScalarExpressionNode(ScalarExpressionNodeKind kind, ScalarExpressionNode? left = null, ScalarExpressionNode? right = null,
                string? text = null, string? errorLiteral = null, string? @operator = null, double factor = 0) {
                Kind = kind;
                Left = left;
                Right = right;
                Text = text;
                ErrorLiteral = errorLiteral;
                Operator = @operator;
                Factor = factor;
            }

            internal ScalarExpressionNodeKind Kind { get; }
            internal ScalarExpressionNode? Left { get; }
            internal ScalarExpressionNode? Right { get; }
            internal string? Text { get; }
            internal string? ErrorLiteral { get; }
            internal string? Operator { get; }
            internal double Factor { get; }

            internal static ScalarExpressionNode Level(ScalarExpressionNode inner) => new(ScalarExpressionNodeKind.Level, inner);
            internal static ScalarExpressionNode Atom(string text) => new(ScalarExpressionNodeKind.Atom, text: text);
            internal static ScalarExpressionNode Error(string literal) => new(ScalarExpressionNodeKind.Atom, errorLiteral: literal);
            internal static ScalarExpressionNode Unary(ScalarExpressionNode operand, double factor) => new(ScalarExpressionNodeKind.Unary, operand, factor: factor);
            internal static ScalarExpressionNode Parentheses(ScalarExpressionNode inner) => new(ScalarExpressionNodeKind.Parentheses, inner);
            internal static ScalarExpressionNode Percent(ScalarExpressionNode operand) => new(ScalarExpressionNodeKind.Percent, operand);
            internal static ScalarExpressionNode Binary(ScalarExpressionNode left, string @operator, ScalarExpressionNode right) =>
                new(ScalarExpressionNodeKind.Binary, left, right, @operator: @operator);
        }

        /// <summary>
        /// Compiles an expression with the interpreting parser's grammar, without evaluating anything. It accepts
        /// only text the parser would read to the end, so a compiled expression evaluates its operands in the same
        /// order and stops at the same failure; anything else returns null and the parser runs as before.
        /// </summary>
        private struct ScalarExpressionCompiler {
            private readonly string _text;
            private int _position;
            private int _depth;

            private ScalarExpressionCompiler(string text) {
                _text = text;
                _position = 0;
                _depth = 0;
            }

            internal static ScalarExpressionNode? Compile(string text) {
                var compiler = new ScalarExpressionCompiler(text);
                ScalarExpressionNode? root = compiler.ParseLevel(1);
                if (root == null) return null;
                compiler.SkipWhiteSpace();
                return compiler._position == text.Length ? root : null;
            }

            private ScalarExpressionNode? ParseLevel(int minimumPrecedence) {
                if (++_depth > 128) { _depth--; return null; }
                try {
                    SkipWhiteSpace();
                    if (_position == _text.Length) return null;
                    char first = _text[_position];
                    ScalarExpressionNode? value;
                    if (first == '+' || first == '-') {
                        _position++;
                        ScalarExpressionNode? operand = ParseLevel(7);
                        if (operand == null) return null;
                        value = ScalarExpressionNode.Unary(operand, first == '-' ? -1 : 1);
                    } else if (first == '(') {
                        _position++;
                        ScalarExpressionNode? inner = ParseLevel(1);
                        if (inner == null) return null;
                        SkipWhiteSpace();
                        if (_position == _text.Length || _text[_position++] != ')') return null;
                        value = ScalarExpressionNode.Parentheses(inner);
                    } else {
                        value = ReadAtom();
                        if (value == null) return null;
                    }

                    while (true) {
                        SkipWhiteSpace();
                        if (_position == _text.Length) break;
                        string op = ScalarFormulaParser.ReadScalarOperator(_text, _position);
                        int precedence = ScalarFormulaParser.Precedence(op);
                        if (precedence < minimumPrecedence) break;
                        _position += op.Length;
                        if (op == "%") {
                            value = ScalarExpressionNode.Percent(value);
                            continue;
                        }
                        ScalarExpressionNode? right = ParseLevel(precedence + 1);
                        if (right == null) return null;
                        value = ScalarExpressionNode.Binary(value, op, right);
                    }
                    return ScalarExpressionNode.Level(value);
                } finally { _depth--; }
            }

            private ScalarExpressionNode? ReadAtom() {
                if (!TryScanScalarAtom(_text, ref _position, out int start, out string? errorLiteral)) return null;
                if (errorLiteral != null) return ScalarExpressionNode.Error(errorLiteral);
                // Operand resolution trims its text first; trimming once here hands it the same instance every pass.
                return ScalarExpressionNode.Atom(_text.Substring(start, _position - start).Trim());
            }

            private void SkipWhiteSpace() {
                while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++;
            }
        }

        private struct ScalarEvaluationState {
            internal bool LastOperationWasAddition;
            internal double LastLeft;
            internal double LastRight;
        }

        private bool TryEvaluateCompiledScalarExpression(ScalarExpressionNode root, out FormulaArgumentValue value) {
            var state = new ScalarEvaluationState();
            if (!TryEvaluateScalarNode(root, ref state, out value)) return false;
            if (state.LastOperationWasAddition && value.Number.HasValue) {
                value = ScalarFormulaParser.CancelAdditionNoise(value, state.LastLeft, state.LastRight);
            }
            return !value.IsUnresolvedFormula;
        }

        private bool TryEvaluateScalarNode(ScalarExpressionNode node, ref ScalarEvaluationState state, out FormulaArgumentValue value) {
            switch (node.Kind) {
                case ScalarExpressionNodeKind.Level:
                    value = default;
                    if (!HasSufficientFormulaExecutionStack()) return false;
                    if (!TryEvaluateScalarNode(node.Left!, ref state, out value)) return false;
                    return !value.IsUnresolvedFormula;
                case ScalarExpressionNodeKind.Atom:
                    if (node.ErrorLiteral != null) {
                        value = FormulaArgumentValue.Error(node.ErrorLiteral);
                        return true;
                    }
                    return TryResolveFormulaArgument(node.Text!, out value, allowScalarExpression: false);
                case ScalarExpressionNodeKind.Unary:
                    if (!TryEvaluateScalarNode(node.Left!, ref state, out value)) return false;
                    value = ScalarFormulaParser.NumericUnary(value, node.Factor);
                    state.LastOperationWasAddition = false;
                    return true;
                case ScalarExpressionNodeKind.Parentheses:
                    if (!TryEvaluateScalarNode(node.Left!, ref state, out value)) return false;
                    state.LastOperationWasAddition = false;
                    return true;
                case ScalarExpressionNodeKind.Percent:
                    if (!TryEvaluateScalarNode(node.Left!, ref state, out value)) return false;
                    value = ScalarFormulaParser.NumericUnary(value, 0.01);
                    state.LastOperationWasAddition = false;
                    return true;
                case ScalarExpressionNodeKind.Binary:
                    value = default;
                    if (!TryEvaluateScalarNode(node.Left!, ref state, out FormulaArgumentValue left)
                        || !TryEvaluateScalarNode(node.Right!, ref state, out FormulaArgumentValue right)) return false;
                    return ScalarFormulaParser.ApplyOperator(left, node.Operator!, right, out value,
                        ref state.LastOperationWasAddition, ref state.LastLeft, ref state.LastRight);
                default:
                    value = default;
                    return false;
            }
        }
    }
}
