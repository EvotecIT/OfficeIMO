using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathMarkup {
    // The HTML adapter supplies bounded, namespace-aware XML with DOM annotations.
    // This hook records token presentation in the shared expression model.
    internal static OfficeMathExpression FromMathMl(XElement root, Action<XElement, OfficeMathExpression> tokenParsed) =>
        new MathMlParser(tokenParsed).ParseMathMlElement(root);

    private sealed class MathMlParser {
        private readonly Action<XElement, OfficeMathExpression>? _tokenParsed;
        private static bool? OperatorBoolean(XElement element, string attribute) =>
            bool.TryParse((string?)element.Attribute(attribute), out bool value) ? value : (bool?)null;
        internal MathMlParser(Action<XElement, OfficeMathExpression>? tokenParsed = null) => _tokenParsed = tokenParsed;
        internal OfficeMathExpression ParseMathMlElement(XElement element) {
            OfficeMathExpression expression = ParseElementCore(element);
            if (element.Name.LocalName is "mi" or "mtext" or "mn" or "mo")
                _tokenParsed?.Invoke(element, expression);
            return expression;
        }

        private OfficeMathExpression ParseElementCore(XElement element) {
            string name = element.Name.LocalName.ToLowerInvariant();
            List<XElement> children = element.Elements().Where(item =>
                !string.Equals(item.Name.LocalName, "annotation", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(item.Name.LocalName, "annotation-xml", StringComparison.OrdinalIgnoreCase)).ToList();
            switch (name) {
                case "math":
                case "mstyle":
                    return CollapseRow(children.Select(ParseMathMlElement));
                case "semantics":
                    return children.Count == 0 ? OfficeMath.Text(string.Empty) : ParseMathMlElement(children[0]);
                case "mrow":
                    if (TryParseFunctionRow(children, out OfficeMathExpression? function)) return function!;
                    OfficeMathExpression[] parsedChildren = children.Select(ParseMathMlElement).ToArray();
                    if (TryParseNaryRow(parsedChildren, out OfficeMathExpression? nary)) return nary!;
                    return CollapseRow(parsedChildren);
                case "mtext": return OfficeMath.Text(element.Value);
                case "mi": return OfficeMath.Identifier(element.Value);
                case "mn": return OfficeMath.Number(element.Value);
                case "mo": return OfficeMath.Create(OfficeMathKind.Operator, text: element.Value,
                    stretchy: OperatorBoolean(element, "stretchy"), largeOperator: OperatorBoolean(element, "largeop"));
                case "mfrac":
                    return string.Equals((string?)element.Attribute("bevelled"), "true", StringComparison.OrdinalIgnoreCase)
                        ? OfficeMath.SlashedFraction(ParseRequired(children, 0, name), ParseRequired(children, 1, name))
                        : OfficeMath.Fraction(ParseRequired(children, 0, name), ParseRequired(children, 1, name));
                case "msqrt": return OfficeMath.Radical(CollapseRow(children.Select(ParseMathMlElement)));
                case "mroot": return OfficeMath.Radical(ParseRequired(children, 0, name), ParseRequired(children, 1, name));
                case "msup": return OfficeMath.Superscript(ParseRequired(children, 0, name), ParseRequired(children, 1, name));
                case "msub": return OfficeMath.Subscript(ParseRequired(children, 0, name), ParseRequired(children, 1, name));
                case "msubsup": return OfficeMath.SubSuperscript(ParseRequired(children, 0, name), ParseRequired(children, 1, name), ParseRequired(children, 2, name));
                case "mmultiscripts": return ParseMultiScripts(children, name);
                case "mfenced":
                    string open = (string?)element.Attribute("open") ?? "(";
                    string close = (string?)element.Attribute("close") ?? ")";
                    OfficeMathExpression fencedContent = CollapseRow(children.Select(ParseMathMlElement));
                    if (children.Count == 1 && children[0].Name.LocalName == "mtable" && open == "[" && close == "]" &&
                        fencedContent.Kind == OfficeMathKind.EquationArray) {
                        return OfficeMath.Create(OfficeMathKind.Matrix, children: fencedContent.Children,
                            rowCount: fencedContent.RowCount, columnCount: fencedContent.ColumnCount);
                    }
                    if (children.Count > 1 || element.Attribute("separators") != null) {
                        return OfficeMath.DelimiterList(open, close, (string?)element.Attribute("separators") ?? ",",
                            children.Select(ParseMathMlElement).ToArray());
                    }
                    return OfficeMath.Delimited(fencedContent, open, close);
                case "mtable": return ParseMathMlTable(element);
                case "menclose": return OfficeMath.Box(CollapseRow(children.Select(ParseMathMlElement)));
                case "mphantom": return OfficeMath.Phantom(CollapseRow(children.Select(ParseMathMlElement)));
                case "mover": return ParseOverUnder(children, over: true, both: false, element);
                case "munder": return ParseOverUnder(children, over: false, both: false, element);
                case "munderover": return ParseOverUnder(children, over: true, both: true, element);
                default:
                    if (children.Count > 0) return CollapseRow(children.Select(ParseMathMlElement));
                    return OfficeMath.Text(element.Value);
            }
        }

        private bool TryParseFunctionRow(IReadOnlyList<XElement> children, out OfficeMathExpression? expression) {
            expression = null;
            if (children.Count != 3 || children[0].Name.LocalName != "mi" || children[1].Name.LocalName != "mo" ||
                children[1].Value != "⁡" || children[2].Name.LocalName != "mfenced") return false;
            OfficeMathExpression argument = ParseMathMlElement(children[2]);
            if (argument.Kind == OfficeMathKind.Delimited && argument.Character == "(" && argument.SecondaryCharacter == ")") {
                argument = argument.Children[0];
            }
            expression = OfficeMath.Function(children[0].Value, argument);
            _tokenParsed?.Invoke(children[0], expression);
            return true;
        }

        private static bool TryParseNaryRow(IReadOnlyList<OfficeMathExpression> children, out OfficeMathExpression? expression) {
            expression = null;
            if (children.Count != 2) return false;
            OfficeMathExpression head = children[0];
            OfficeMathExpression content = children[1];
            if (head.Kind == OfficeMathKind.Operator && head.LargeOperator != false && IsNarySymbol(head.Text)) {
                expression = OfficeMath.Nary(head.Text!, content);
                return true;
            }
            if (head.Kind != OfficeMathKind.Nary || head.Children.Count == 0 || head.Children[0].ToPlainText().Length != 0) return false;
            expression = OfficeMath.Nary(head.Character ?? "∑", content, head.NaryLowerLimit, head.NaryUpperLimit);
            return true;
        }

        private OfficeMathExpression ParseOverUnder(List<XElement> children, bool over, bool both, XElement source) {
            OfficeMathExpression basis = ParseRequired(children, 0, source.Name.LocalName);
            OfficeMathExpression first = ParseRequired(children, 1, source.Name.LocalName);
            if (basis.Kind == OfficeMathKind.Operator && basis.LargeOperator != false && IsNarySymbol(basis.Text)) {
                OfficeMathExpression content = OfficeMath.Text(string.Empty);
                return both
                    ? OfficeMath.Nary(basis.Text!, content, first, ParseRequired(children, 2, source.Name.LocalName))
                    : over ? OfficeMath.Nary(basis.Text!, content, null, first) : OfficeMath.Nary(basis.Text!, content, first);
            }
            if (both) return OfficeMath.SubSuperscript(basis, first, ParseRequired(children, 2, source.Name.LocalName));
            bool accent = string.Equals((string?)source.Attribute(over ? "accent" : "accentunder"), "true", StringComparison.OrdinalIgnoreCase);
            if (accent && first.Stretchy != false && first.ToPlainText() == (over ? "¯" : "_")) return over ? OfficeMath.Overbar(basis) : OfficeMath.Underbar(basis);
            if (accent && over) {
                OfficeMathExpression accented = OfficeMath.Create(OfficeMathKind.Accent, children: new[] { basis },
                    character: first.ToPlainText(), stretchy: first.Stretchy);
                _tokenParsed?.Invoke(children[1], accented);
                return accented;
            }
            return over ? OfficeMath.UpperLimit(basis, first) : OfficeMath.LowerLimit(basis, first);
        }

        private OfficeMathExpression ParseMultiScripts(IReadOnlyList<XElement> children, string owner) {
            int marker = -1;
            for (int index = 0; index < children.Count; index++) {
                if (children[index].Name.LocalName == "mprescripts") { marker = index; break; }
            }
            if (marker < 0) {
                if (children.Count != 3) throw new FormatException("MathML element '" + owner + "' has invalid postscripts.");
                return ApplyPostScripts(ParseMathMlElement(children[0]), children[1], children[2]);
            }
            if ((marker != 1 && marker != 3) || marker + 3 != children.Count) {
                throw new FormatException("MathML element '" + owner + "' has invalid script pairs.");
            }
            OfficeMathExpression basis = ParseMathMlElement(children[0]);
            if (marker == 3) basis = ApplyPostScripts(basis, children[1], children[2]);
            return OfficeMath.LeftSubSuperscript(
                basis,
                ParseMathMlElement(children[marker + 1]),
                ParseMathMlElement(children[marker + 2]));
        }

        private OfficeMathExpression ApplyPostScripts(
            OfficeMathExpression basis,
            XElement subscript,
            XElement superscript) {
            bool hasSubscript = subscript.Name.LocalName != "none";
            bool hasSuperscript = superscript.Name.LocalName != "none";
            if (hasSubscript && hasSuperscript) {
                return OfficeMath.SubSuperscript(basis, ParseMathMlElement(subscript), ParseMathMlElement(superscript));
            }
            if (hasSubscript) return OfficeMath.Subscript(basis, ParseMathMlElement(subscript));
            if (hasSuperscript) return OfficeMath.Superscript(basis, ParseMathMlElement(superscript));
            return basis;
        }

        private OfficeMathExpression ParseMathMlTable(XElement table) {
            List<XElement> rows = table.Elements().Where(item => item.Name.LocalName == "mtr" || item.Name.LocalName == "mlabeledtr").ToList();
            if (rows.Count == 0) return OfficeMath.EquationArray(1, 1, OfficeMath.Text(string.Empty));
            int columns = Math.Max(1, rows.Max(row => row.Elements().Count(item => item.Name.LocalName == "mtd")));
            var cells = new List<OfficeMathExpression>(rows.Count * columns);
            foreach (XElement row in rows) {
                List<XElement> rowCells = row.Elements().Where(item => item.Name.LocalName == "mtd").ToList();
                for (int column = 0; column < columns; column++) {
                    cells.Add(column < rowCells.Count
                        ? CollapseRow(rowCells[column].Elements().Select(ParseMathMlElement))
                        : OfficeMath.Text(string.Empty));
                }
            }
            string? kind = (string?)table.Attribute("data-officeimo-kind");
            if (kind == "stack" || kind == "stretch-stack") {
                OfficeMathExpression[] stackRows = rows.Select(row => CollapseRow(
                    row.Elements().Where(item => item.Name.LocalName == "mtd").SelectMany(cell => cell.Elements()).Select(ParseMathMlElement))).ToArray();
                return kind == "stretch-stack" ? OfficeMath.StretchStack(stackRows) : OfficeMath.Stack(stackRows);
            }
            return OfficeMath.EquationArray(rows.Count, columns, cells.ToArray());
        }

        private OfficeMathExpression ParseRequired(IReadOnlyList<XElement> children, int index, string owner) {
            if (index >= children.Count) throw new FormatException("MathML element '" + owner + "' has too few operands.");
            return ParseMathMlElement(children[index]);
        }

    }
}
