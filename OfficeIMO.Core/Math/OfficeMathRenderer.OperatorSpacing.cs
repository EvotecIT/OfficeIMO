using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        // A root token is inside the math element's implicit row. Nested rows apply
        // their own spacing; a child's raw Layout deliberately excludes outside space.
        internal LayoutBox LayoutRoot(OfficeMathExpression expression, double scale) =>
            AddOperatorSpace(Layout(expression, scale), expression, "infix", scale);

        private LayoutBox AddOperatorSpace(LayoutBox box, OfficeMathExpression expression, string inferredForm, double scale) {
            if (expression.Kind != OfficeMathKind.Operator) return box;
            string form = expression.OperatorForm ?? inferredForm;
            bool known = TryOperatorDictionary(expression.Text, form, out double leftEm, out double rightEm);
            if (!known && expression.OperatorForm == null) {
                // MathML Core tries the other forms only when form was not authored.
                known = TryOperatorDictionary(expression.Text, "infix", out leftEm, out rightEm)
                    || TryOperatorDictionary(expression.Text, "postfix", out leftEm, out rightEm)
                    || TryOperatorDictionary(expression.Text, "prefix", out leftEm, out rightEm);
            }
            if (!known) leftEm = rightEm = 5D / 18D;
            double em = FontSize(scale);
            double pixels = _options.MathMlPixelScale ?? _options.Dpi / 96D;
            double left = OfficeMathOperatorSpacing.ResolveLength(expression.OperatorLeftSpace, leftEm * em, em, pixels);
            double right = OfficeMathOperatorSpacing.ResolveLength(expression.OperatorRightSpace, rightEm * em, em, pixels);
            if (left == 0D && right == 0D) return box;
            double width = box.Width + left + right;
            var spaced = new LayoutBox(width, box.Height, box.Baseline);
            spaced.Add(box, left, 0D);
            spaced.ItalicCorrection = box.ItalicCorrection;
            return spaced;
        }

        private static void FindNonSpaceRowChildren(IReadOnlyList<OfficeMathExpression> children, out int first, out int last) {
            first = last = -1;
            for (int i = 0; i < children.Count; i++) {
                if (IsMathSpaceLike(children[i])) continue;
                if (first < 0) first = i;
                last = i;
            }
        }

        private static bool IsMathSpaceLike(OfficeMathExpression expression) {
            if (expression.Kind == OfficeMathKind.Text) return true;
            if (expression.Kind is not (OfficeMathKind.Row or OfficeMathKind.Phantom)) return false;
            for (int i = 0; i < expression.Children.Count; i++)
                if (!IsMathSpaceLike(expression.Children[i])) return false;
            return true;
        }

        // Selected ordinary arithmetic, relation, fence and punctuation entries from
        // MathML Core 2025-06-24 Appendix B. This is not the complete operator dictionary.
        private static bool TryOperatorDictionary(string? text, string form, out double left, out double right) {
            left = right = 0D;
            if (form == "infix") {
                switch (text) {
                    case "+": case "-": case "−": case "±": case "÷": left = right = 2D / 9D; return true;
                    case "×": case "*": case "·": left = right = 1D / 6D; return true;
                    case "=": case "<": case ">": case "≤": case "≥": case "≠": case "|":
                        left = right = 5D / 18D; return true;
                    case ",": case ":": case ";": right = 1D / 6D; return true;
                    case "/": case "\\": case "⁡": case "⁢": case "⁣": case "⁤": return true;
                }
            } else if (form == "prefix") {
                switch (text) {
                    case "+": case "-": case "−": case "±": case "!":
                    case "(": case "[": case "{": case "|": case "‖": case "⟨": return true;
                    case "∑": case "∏": case "∫": case "⋂": case "⋃": left = right = 1D / 6D; return true;
                }
            } else if (form == "postfix") {
                switch (text) {
                    case "!": case "'": case "′": case "\"": case "″": case "%":
                    case ")": case "]": case "}": case "|": case "‖": case "⟩": return true;
                }
            }
            return false;
        }
    }
}
