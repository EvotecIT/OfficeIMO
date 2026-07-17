using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

/// <summary>MathML and LaTeX conversion for the shared mathematical expression tree.</summary>
public static class OfficeMathMarkup {
    private static readonly XNamespace MathMlNamespace = "http://www.w3.org/1998/Math/MathML";
    /// <summary>Default maximum nested structure depth accepted by portable math parsers.</summary>
    public const int DefaultMaximumParseDepth = 128;

    /// <summary>Serializes an expression as presentation MathML.</summary>
    public static string ToMathMl(OfficeMathExpression expression, bool display = false) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        var root = new XElement(MathMlNamespace + "math", ToMathMlElement(expression));
        if (display) root.SetAttributeValue("display", "block");
        return root.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Parses presentation MathML into the shared expression tree.</summary>
    public static OfficeMathExpression FromMathMl(string mathMl) => FromMathMl(mathMl, DefaultMaximumParseDepth);

    /// <summary>Parses presentation MathML with a hard nesting-depth limit.</summary>
    public static OfficeMathExpression FromMathMl(string mathMl, int maximumDepth) {
        if (string.IsNullOrWhiteSpace(mathMl)) throw new ArgumentException("MathML is required.", nameof(mathMl));
        if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        ValidateXmlDepth(mathMl, maximumDepth);
        XDocument document;
        try {
            document = XDocument.Parse(mathMl, LoadOptions.None);
        } catch (Exception exception) when (exception is System.Xml.XmlException || exception is ArgumentException) {
            throw new FormatException("The MathML is not well formed.", exception);
        }
        if (document.Root == null) throw new FormatException("The MathML has no root element.");
        return ParseMathMlElement(document.Root);
    }

    /// <summary>Serializes an expression as portable LaTeX math source.</summary>
    public static string ToLatex(OfficeMathExpression expression) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        var builder = new StringBuilder();
        AppendLatex(builder, expression);
        return builder.ToString();
    }

    /// <summary>Parses common LaTeX math constructs without requiring a TeX runtime.</summary>
    public static OfficeMathExpression FromLatex(string latex) => FromLatex(latex, DefaultMaximumParseDepth);

    /// <summary>Parses common LaTeX math constructs with a hard nesting-depth limit.</summary>
    public static OfficeMathExpression FromLatex(string latex, int maximumDepth) {
        if (latex == null) throw new ArgumentNullException(nameof(latex));
        if (maximumDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        return new LatexParser(latex, maximumDepth).Parse();
    }

    internal static void ValidateXmlDepth(string mathMl, int maximumDepth) {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        try {
            using var text = new StringReader(mathMl);
            using XmlReader reader = XmlReader.Create(text, settings);
            while (reader.Read()) {
                if (reader.Depth >= maximumDepth) {
                    throw new OfficeMathParseException("DRAWING_MATH_DEPTH", "The mathematical markup nesting depth limit was exceeded.");
                }
            }
        } catch (OfficeMathParseException) {
            throw;
        } catch (XmlException exception) {
            throw new FormatException("The MathML is not well formed.", exception);
        }
    }

    private static XElement ToMathMlElement(OfficeMathExpression expression) {
        switch (expression.Kind) {
            case OfficeMathKind.Text: return Element("mtext", expression.Text);
            case OfficeMathKind.Identifier: return Element("mi", expression.Text);
            case OfficeMathKind.Number: return Element("mn", expression.Text);
            case OfficeMathKind.Operator: return Element("mo", expression.Text);
            case OfficeMathKind.Row: return Element("mrow", expression.Children.Select(ToMathMlElement));
            case OfficeMathKind.Fraction:
                return Element("mfrac", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.SlashedFraction:
                return new XElement(MathMlNamespace + "mfrac", new XAttribute("bevelled", "true"),
                    ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.Radical:
                return expression.Children.Count == 1
                    ? Element("msqrt", ToMathMlElement(expression.Children[0]))
                    : Element("mroot", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.Superscript:
                return Element("msup", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.Subscript:
                return Element("msub", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.SubSuperscript:
                return Element("msubsup", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]), ToMathMlElement(expression.Children[2]));
            case OfficeMathKind.LeftSubSuperscript:
                return Element("mmultiscripts", ToMathMlElement(expression.Children[0]),
                    new XElement(MathMlNamespace + "mprescripts"),
                    ToMathMlElement(expression.Children[1]), ToMathMlElement(expression.Children[2]));
            case OfficeMathKind.LowerLimit:
                return Element("munder", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.UpperLimit:
                return Element("mover", ToMathMlElement(expression.Children[0]), ToMathMlElement(expression.Children[1]));
            case OfficeMathKind.Delimited:
                return new XElement(MathMlNamespace + "mfenced",
                    new XAttribute("open", expression.Character ?? "("),
                    new XAttribute("close", expression.SecondaryCharacter ?? ")"),
                    ToMathMlElement(expression.Children[0]));
            case OfficeMathKind.DelimiterList:
                return new XElement(MathMlNamespace + "mfenced",
                    new XAttribute("open", expression.Character ?? "("),
                    new XAttribute("close", expression.SecondaryCharacter ?? ")"),
                    new XAttribute("separators", expression.SeparatorCharacter ?? ","),
                    expression.Children.Select(ToMathMlElement));
            case OfficeMathKind.Function:
                return Element("mrow", Element("mi", expression.Text), new XElement(MathMlNamespace + "mo", "⁡"),
                    new XElement(MathMlNamespace + "mfenced", ToMathMlElement(expression.Children[0])));
            case OfficeMathKind.Nary:
                return NaryMathMl(expression);
            case OfficeMathKind.Matrix:
            case OfficeMathKind.EquationArray:
                return MatrixMathMl(expression);
            case OfficeMathKind.Accent:
                return new XElement(MathMlNamespace + "mover", new XAttribute("accent", "true"),
                    ToMathMlElement(expression.Children[0]), Element("mo", expression.Character ?? "^"));
            case OfficeMathKind.Overbar:
                return new XElement(MathMlNamespace + "mover", new XAttribute("accent", "true"),
                    ToMathMlElement(expression.Children[0]), Element("mo", "¯"));
            case OfficeMathKind.Underbar:
                return new XElement(MathMlNamespace + "munder", new XAttribute("accentunder", "true"),
                    ToMathMlElement(expression.Children[0]), Element("mo", "_"));
            case OfficeMathKind.Box:
                return new XElement(MathMlNamespace + "menclose", new XAttribute("notation", "box"), ToMathMlElement(expression.Children[0]));
            case OfficeMathKind.Phantom:
                return Element("mphantom", ToMathMlElement(expression.Children[0]));
            case OfficeMathKind.Stack:
            case OfficeMathKind.StretchStack:
                return StackMathMl(expression);
            default:
                throw new NotSupportedException("Unsupported mathematical node: " + expression.Kind + ".");
        }
    }

    private static XElement NaryMathMl(OfficeMathExpression expression) {
        XElement symbol = Element("mo", expression.Character ?? "∑");
        XElement withLimits = symbol;
        if (expression.Children.Count == 2) {
            withLimits = Element("munder", symbol, ToMathMlElement(expression.Children[1]));
        } else if (expression.Children.Count == 3) {
            withLimits = Element("munderover", symbol, ToMathMlElement(expression.Children[1]), ToMathMlElement(expression.Children[2]));
        }
        return Element("mrow", withLimits, ToMathMlElement(expression.Children[0]));
    }

    private static XElement MatrixMathMl(OfficeMathExpression expression) {
        var table = new XElement(MathMlNamespace + "mtable");
        for (int row = 0; row < expression.RowCount; row++) {
            var tableRow = new XElement(MathMlNamespace + "mtr");
            for (int column = 0; column < expression.ColumnCount; column++) {
                tableRow.Add(Element("mtd", ToMathMlElement(expression.Children[row * expression.ColumnCount + column])));
            }
            table.Add(tableRow);
        }
        return expression.Kind == OfficeMathKind.Matrix
            ? new XElement(MathMlNamespace + "mfenced", new XAttribute("open", "["), new XAttribute("close", "]"), table)
            : table;
    }

    private static XElement StackMathMl(OfficeMathExpression expression) {
        var table = new XElement(MathMlNamespace + "mtable",
            new XAttribute("data-officeimo-kind", expression.Kind == OfficeMathKind.StretchStack ? "stretch-stack" : "stack"));
        foreach (OfficeMathExpression row in expression.Children) {
            table.Add(Element("mtr", Element("mtd", ToMathMlElement(row))));
        }
        return table;
    }

    private static OfficeMathExpression ParseMathMlElement(XElement element) {
        string name = element.Name.LocalName.ToLowerInvariant();
        List<XElement> children = element.Elements().Where(item => item.Name.LocalName != "annotation").ToList();
        switch (name) {
            case "math":
            case "mstyle":
            case "semantics":
                return CollapseRow(children.Select(ParseMathMlElement));
            case "mrow":
                if (TryParseFunctionRow(children, out OfficeMathExpression? function)) return function!;
                if (TryParseNaryRow(children, out OfficeMathExpression? nary)) return nary!;
                return CollapseRow(children.Select(ParseMathMlElement));
            case "mtext": return OfficeMath.Text(element.Value);
            case "mi": return OfficeMath.Identifier(element.Value);
            case "mn": return OfficeMath.Number(element.Value);
            case "mo": return OfficeMath.Operator(element.Value);
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

    private static bool TryParseFunctionRow(IReadOnlyList<XElement> children, out OfficeMathExpression? expression) {
        expression = null;
        if (children.Count != 3 || children[0].Name.LocalName != "mi" || children[1].Name.LocalName != "mo" ||
            children[1].Value != "⁡" || children[2].Name.LocalName != "mfenced") return false;
        OfficeMathExpression argument = ParseMathMlElement(children[2]);
        if (argument.Kind == OfficeMathKind.Delimited && argument.Character == "(" && argument.SecondaryCharacter == ")") {
            argument = argument.Children[0];
        }
        expression = OfficeMath.Function(children[0].Value, argument);
        return true;
    }

    private static bool TryParseNaryRow(IReadOnlyList<XElement> children, out OfficeMathExpression? expression) {
        expression = null;
        if (children.Count != 2) return false;
        OfficeMathExpression head = ParseMathMlElement(children[0]);
        OfficeMathExpression content = ParseMathMlElement(children[1]);
        if (head.Kind == OfficeMathKind.Operator && IsNarySymbol(head.Text)) {
            expression = OfficeMath.Nary(head.Text!, content);
            return true;
        }
        if (head.Kind != OfficeMathKind.Nary || head.Children.Count == 0 || head.Children[0].ToPlainText().Length != 0) return false;
        OfficeMathExpression? lower = head.Children.Count > 1 ? head.Children[1] : null;
        OfficeMathExpression? upper = head.Children.Count > 2 ? head.Children[2] : null;
        expression = OfficeMath.Nary(head.Character ?? "∑", content, lower, upper);
        return true;
    }

    private static OfficeMathExpression ParseOverUnder(List<XElement> children, bool over, bool both, XElement source) {
        OfficeMathExpression basis = ParseRequired(children, 0, source.Name.LocalName);
        OfficeMathExpression first = ParseRequired(children, 1, source.Name.LocalName);
        if (basis.Kind == OfficeMathKind.Operator && IsNarySymbol(basis.Text)) {
            OfficeMathExpression content = OfficeMath.Text(string.Empty);
            return both
                ? OfficeMath.Nary(basis.Text!, content, first, ParseRequired(children, 2, source.Name.LocalName))
                : over ? OfficeMath.Nary(basis.Text!, content, null, first) : OfficeMath.Nary(basis.Text!, content, first);
        }
        if (both) return OfficeMath.SubSuperscript(basis, first, ParseRequired(children, 2, source.Name.LocalName));
        bool accent = string.Equals((string?)source.Attribute(over ? "accent" : "accentunder"), "true", StringComparison.OrdinalIgnoreCase);
        if (accent && first.ToPlainText() == (over ? "¯" : "_")) return over ? OfficeMath.Overbar(basis) : OfficeMath.Underbar(basis);
        if (accent && over) return OfficeMath.Accent(basis, first.ToPlainText());
        return over ? OfficeMath.UpperLimit(basis, first) : OfficeMath.LowerLimit(basis, first);
    }

    private static OfficeMathExpression ParseMultiScripts(IReadOnlyList<XElement> children, string owner) {
        int marker = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index].Name.LocalName == "mprescripts") { marker = index; break; }
        }
        if (marker < 1 || marker + 2 >= children.Count) throw new FormatException("MathML element '" + owner + "' has invalid prescripts.");
        return OfficeMath.LeftSubSuperscript(
            ParseMathMlElement(children[0]),
            ParseMathMlElement(children[marker + 1]),
            ParseMathMlElement(children[marker + 2]));
    }

    private static OfficeMathExpression ParseMathMlTable(XElement table) {
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

    private static OfficeMathExpression ParseRequired(IReadOnlyList<XElement> children, int index, string owner) {
        if (index >= children.Count) throw new FormatException("MathML element '" + owner + "' has too few operands.");
        return ParseMathMlElement(children[index]);
    }

    private static OfficeMathExpression CollapseRow(IEnumerable<OfficeMathExpression> expressions) {
        OfficeMathExpression[] values = expressions.ToArray();
        if (values.Length == 0) return OfficeMath.Text(string.Empty);
        return values.Length == 1 ? values[0] : OfficeMath.Row(values);
    }

    private static XElement Element(string name, params object?[] content) => new XElement(MathMlNamespace + name, content);

    private static void AppendLatex(StringBuilder builder, OfficeMathExpression expression) {
        switch (expression.Kind) {
            case OfficeMathKind.Text: builder.Append("\\text{").Append(EscapeLatexText(expression.Text ?? string.Empty)).Append('}'); break;
            case OfficeMathKind.Identifier:
            case OfficeMathKind.Number: builder.Append(expression.Text); break;
            case OfficeMathKind.Operator: builder.Append(SymbolToLatex(expression.Text)); break;
            case OfficeMathKind.Row:
                foreach (OfficeMathExpression child in expression.Children) AppendLatex(builder, child);
                break;
            case OfficeMathKind.Fraction:
                builder.Append("\\frac{"); AppendLatex(builder, expression.Children[0]); builder.Append("}{"); AppendLatex(builder, expression.Children[1]); builder.Append('}'); break;
            case OfficeMathKind.SlashedFraction:
                builder.Append("\\sfrac{"); AppendLatex(builder, expression.Children[0]); builder.Append("}{"); AppendLatex(builder, expression.Children[1]); builder.Append('}'); break;
            case OfficeMathKind.Radical:
                builder.Append("\\sqrt");
                if (expression.Children.Count == 2) { builder.Append('['); AppendLatex(builder, expression.Children[1]); builder.Append(']'); }
                builder.Append('{'); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Superscript:
                AppendGroupedBase(builder, expression.Children[0]); builder.Append("^{"); AppendLatex(builder, expression.Children[1]); builder.Append('}'); break;
            case OfficeMathKind.Subscript:
                AppendGroupedBase(builder, expression.Children[0]); builder.Append("_{"); AppendLatex(builder, expression.Children[1]); builder.Append('}'); break;
            case OfficeMathKind.SubSuperscript:
                AppendGroupedBase(builder, expression.Children[0]); builder.Append("_{"); AppendLatex(builder, expression.Children[1]); builder.Append("}^{"); AppendLatex(builder, expression.Children[2]); builder.Append('}'); break;
            case OfficeMathKind.LeftSubSuperscript:
                builder.Append("\\prescript{"); AppendLatex(builder, expression.Children[2]); builder.Append("}{"); AppendLatex(builder, expression.Children[1]); builder.Append("}{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.LowerLimit:
                builder.Append("\\underset{"); AppendLatex(builder, expression.Children[1]); builder.Append("}{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.UpperLimit:
                builder.Append("\\overset{"); AppendLatex(builder, expression.Children[1]); builder.Append("}{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Delimited:
                builder.Append("\\left").Append(DelimiterToLatex(expression.Character)); AppendLatex(builder, expression.Children[0]); builder.Append("\\right").Append(DelimiterToLatex(expression.SecondaryCharacter)); break;
            case OfficeMathKind.DelimiterList:
                builder.Append("\\delimiterlist{").Append(EscapeLatexText(expression.Character ?? "(")).Append("}{")
                    .Append(EscapeLatexText(expression.SecondaryCharacter ?? ")")).Append("}{")
                    .Append(EscapeLatexText(expression.SeparatorCharacter ?? ",")).Append("}{")
                    .Append(expression.Children.Count.ToString(CultureInfo.InvariantCulture)).Append('}');
                foreach (OfficeMathExpression child in expression.Children) {
                    builder.Append('{'); AppendLatex(builder, child); builder.Append('}');
                }
                break;
            case OfficeMathKind.Function:
                builder.Append('\\').Append(expression.Text).Append("\\left("); AppendLatex(builder, expression.Children[0]); builder.Append("\\right)"); break;
            case OfficeMathKind.Nary:
                builder.Append(SymbolToLatex(expression.Character));
                if (expression.Children.Count > 1) { builder.Append("_{"); AppendLatex(builder, expression.Children[1]); builder.Append('}'); }
                if (expression.Children.Count > 2) { builder.Append("^{"); AppendLatex(builder, expression.Children[2]); builder.Append('}'); }
                builder.Append(" {"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Matrix:
            case OfficeMathKind.EquationArray:
                AppendLatexMatrix(builder, expression); break;
            case OfficeMathKind.Accent:
                builder.Append("\\accent{").Append(EscapeLatexText(expression.Character ?? "^")).Append("}{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Overbar:
                builder.Append("\\overline{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Underbar:
                builder.Append("\\underline{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Box:
                builder.Append("\\boxed{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Phantom:
                builder.Append("\\phantom{"); AppendLatex(builder, expression.Children[0]); builder.Append('}'); break;
            case OfficeMathKind.Stack:
            case OfficeMathKind.StretchStack:
                AppendLatexStack(builder, expression); break;
        }
    }

    private static void AppendLatexMatrix(StringBuilder builder, OfficeMathExpression expression) {
        string environment = expression.Kind == OfficeMathKind.Matrix ? "bmatrix" : "aligned";
        builder.Append("\\begin{").Append(environment).Append('}');
        for (int row = 0; row < expression.RowCount; row++) {
            if (row > 0) builder.Append("\\\\");
            for (int column = 0; column < expression.ColumnCount; column++) {
                if (column > 0) builder.Append('&');
                AppendLatex(builder, expression.Children[row * expression.ColumnCount + column]);
            }
        }
        builder.Append("\\end{").Append(environment).Append('}');
    }

    private static void AppendLatexStack(StringBuilder builder, OfficeMathExpression expression) {
        string environment = expression.Kind == OfficeMathKind.StretchStack ? "officeimostretchstack" : "gathered";
        builder.Append("\\begin{").Append(environment).Append('}');
        for (int index = 0; index < expression.Children.Count; index++) {
            if (index > 0) builder.Append("\\\\");
            AppendLatex(builder, expression.Children[index]);
        }
        builder.Append("\\end{").Append(environment).Append('}');
    }

    private static void AppendGroupedBase(StringBuilder builder, OfficeMathExpression expression) {
        bool group = expression.Kind == OfficeMathKind.Row || expression.Kind == OfficeMathKind.Fraction;
        if (group) builder.Append('{');
        AppendLatex(builder, expression);
        if (group) builder.Append('}');
    }

    private static string EscapeLatexText(string text) => text.Replace("\\", "\\backslash ").Replace("{", "\\{").Replace("}", "\\}").Replace("#", "\\#").Replace("%", "\\%").Replace("&", "\\&").Replace("_", "\\_").Replace("^", "\\^");

    private static string SymbolToLatex(string? symbol) {
        switch (symbol) {
            case "∑": return "\\sum";
            case "∏": return "\\prod";
            case "∫": return "\\int";
            case "∞": return "\\infty";
            case "≤": return "\\le";
            case "≥": return "\\ge";
            case "≠": return "\\ne";
            case "×": return "\\times";
            case "÷": return "\\div";
            case "±": return "\\pm";
            case "→": return "\\to";
            default: return symbol ?? string.Empty;
        }
    }

    private static string DelimiterToLatex(string? delimiter) => delimiter == "{" ? "\\{" : delimiter == "}" ? "\\}" : delimiter ?? ".";

    private static bool IsNarySymbol(string? value) => value == "∑" || value == "∏" || value == "∫" || value == "⋂" || value == "⋃";

    private sealed class LatexParser {
        private static readonly Dictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["sum"] = "∑", ["prod"] = "∏", ["int"] = "∫", ["infty"] = "∞", ["le"] = "≤", ["leq"] = "≤",
            ["ge"] = "≥", ["geq"] = "≥", ["ne"] = "≠", ["neq"] = "≠", ["times"] = "×", ["div"] = "÷",
            ["pm"] = "±", ["to"] = "→", ["alpha"] = "α", ["beta"] = "β", ["gamma"] = "γ", ["delta"] = "δ",
            ["theta"] = "θ", ["lambda"] = "λ", ["mu"] = "μ", ["pi"] = "π", ["sigma"] = "σ", ["phi"] = "φ", ["omega"] = "ω"
        };
        private static readonly HashSet<string> Functions = new HashSet<string>(new[] { "sin", "cos", "tan", "log", "ln", "exp", "lim", "min", "max" }, StringComparer.Ordinal);
        private readonly string _text;
        private readonly int _maximumDepth;
        private int _depth;
        private int _position;

        internal LatexParser(string text, int maximumDepth, int initialDepth = 0) {
            _text = text;
            _maximumDepth = maximumDepth;
            _depth = initialDepth;
        }

        internal OfficeMathExpression Parse() {
            OfficeMathExpression expression = ParseSequence('\0');
            SkipWhitespace();
            if (_position != _text.Length) throw Error("Unexpected trailing input.");
            return expression;
        }

        private OfficeMathExpression ParseSequence(char terminator) {
            if (++_depth > _maximumDepth) {
                _depth--;
                throw new OfficeMathParseException("DRAWING_MATH_DEPTH", "The mathematical markup nesting depth limit was exceeded.");
            }
            try {
                var items = new List<OfficeMathExpression>();
                while (_position < _text.Length) {
                    SkipWhitespace();
                    if (_position >= _text.Length || (terminator != '\0' && _text[_position] == terminator)) break;
                    OfficeMathExpression atom = ParseAtom();
                    OfficeMathExpression? subscript = null;
                    OfficeMathExpression? superscript = null;
                    while (true) {
                        SkipWhitespace();
                        if (_position >= _text.Length || (_text[_position] != '_' && _text[_position] != '^')) break;
                        char marker = _text[_position++];
                        OfficeMathExpression script = ParseScript();
                        if (marker == '_') subscript = script; else superscript = script;
                    }
                    if (IsNaryExpression(atom)) {
                        SkipWhitespace();
                        OfficeMathExpression content = _position < _text.Length &&
                            (terminator == '\0' || _text[_position] != terminator)
                            ? (Peek('{') ? ParseRequiredGroup() : ParseAtom())
                            : OfficeMath.Text(string.Empty);
                        atom = OfficeMath.Nary(atom.Text!, content, subscript, superscript);
                    } else if (subscript != null && superscript != null) atom = OfficeMath.SubSuperscript(atom, subscript, superscript);
                    else if (subscript != null) atom = OfficeMath.Subscript(atom, subscript);
                    else if (superscript != null) atom = OfficeMath.Superscript(atom, superscript);
                    items.Add(atom);
                }
                return CollapseRow(items);
            } finally {
                _depth--;
            }
        }

        private OfficeMathExpression ParseAtom() {
            if (_position >= _text.Length) return OfficeMath.Text(string.Empty);
            char value = _text[_position++];
            if (value == '{') {
                OfficeMathExpression group = ParseSequence('}');
                Require('}');
                return group;
            }
            if (value == '}') throw Error("Unexpected closing brace.");
            if (value == '\\') return ParseCommand();
            if (char.IsDigit(value) || value == '.') return OfficeMath.Number(ReadWhile(value, character => char.IsDigit(character) || character == '.'));
            if (char.IsLetter(value)) return OfficeMath.Identifier(ReadWhile(value, char.IsLetterOrDigit));
            return OfficeMath.Operator(value.ToString(CultureInfo.InvariantCulture));
        }

        private OfficeMathExpression ParseCommand() {
            string command = ReadCommandName();
            switch (command) {
                case "frac": return OfficeMath.Fraction(ParseRequiredGroup(), ParseRequiredGroup());
                case "sfrac": return OfficeMath.SlashedFraction(ParseRequiredGroup(), ParseRequiredGroup());
                case "sqrt":
                    OfficeMathExpression? index = null;
                    SkipWhitespace();
                    if (Peek('[')) { _position++; index = ParseSequence(']'); Require(']'); }
                    OfficeMathExpression radicand = ParseRequiredGroup();
                    return index == null ? OfficeMath.Radical(radicand) : OfficeMath.Radical(radicand, index);
                case "overline": return OfficeMath.Overbar(ParseRequiredGroup());
                case "underline": return OfficeMath.Underbar(ParseRequiredGroup());
                case "boxed": return OfficeMath.Box(ParseRequiredGroup());
                case "phantom": return OfficeMath.Phantom(ParseRequiredGroup());
                case "text": return ParseTextGroup();
                case "backslash": return OfficeMath.Text("\\");
                case "_":
                case "^":
                case "{":
                case "}":
                case "#":
                case "%":
                case "&": return OfficeMath.Text(command);
                case "prescript":
                    OfficeMathExpression leftSup = ParseRequiredGroup();
                    OfficeMathExpression leftSub = ParseRequiredGroup();
                    return OfficeMath.LeftSubSuperscript(ParseRequiredGroup(), leftSub, leftSup);
                case "delimiterlist":
                    string listLeft = ParseRequiredGroup().ToPlainText();
                    string listRight = ParseRequiredGroup().ToPlainText();
                    string listSeparator = ParseRequiredGroup().ToPlainText();
                    string countText = ParseRequiredGroup().ToPlainText();
                    if (!int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out int itemCount) || itemCount < 1 || itemCount > 4096) {
                        throw Error("A delimiter list has an invalid item count.");
                    }
                    var listItems = new OfficeMathExpression[itemCount];
                    for (int itemIndex = 0; itemIndex < itemCount; itemIndex++) listItems[itemIndex] = ParseRequiredGroup();
                    return OfficeMath.DelimiterList(listLeft, listRight, listSeparator, listItems);
                case "underset":
                    OfficeMathExpression lower = ParseRequiredGroup();
                    return OfficeMath.LowerLimit(ParseRequiredGroup(), lower);
                case "overset":
                    OfficeMathExpression upper = ParseRequiredGroup();
                    return OfficeMath.UpperLimit(ParseRequiredGroup(), upper);
                case "accent":
                    OfficeMathExpression accent = ParseRequiredGroup();
                    return OfficeMath.Accent(ParseRequiredGroup(), accent.ToPlainText());
                case "begin": return ParseEnvironment();
                case "left":
                    string left = ReadDelimiter();
                    int rightIndex = FindMatchingRight();
                    if (rightIndex < 0) throw Error("A \\left delimiter has no matching \\right.");
                    string inner = _text.Substring(_position, rightIndex - _position);
                    _position = rightIndex + 6;
                    string right = ReadDelimiter();
                    return OfficeMath.Delimited(new LatexParser(inner, _maximumDepth, _depth).Parse(), left, right);
                default:
                    if (Symbols.TryGetValue(command, out string? symbol)) return OfficeMath.Operator(symbol);
                    if (Functions.Contains(command)) {
                        return OfficeMath.Function(command, ParseFunctionArgument());
                    }
                    if (NextIsLeftDelimiter()) return OfficeMath.Function(command, ParseFunctionArgument());
                    return OfficeMath.Identifier(command);
            }
        }

        private OfficeMathExpression ParseFunctionArgument() {
            SkipWhitespace();
            if (_position >= _text.Length) return OfficeMath.Text(string.Empty);
            OfficeMathExpression argument = Peek('{') ? ParseRequiredGroup() : ParseAtom();
            return argument.Kind == OfficeMathKind.Delimited && argument.Character == "(" && argument.SecondaryCharacter == ")"
                ? argument.Children[0]
                : argument;
        }

        private bool NextIsLeftDelimiter() {
            SkipWhitespace();
            return _position + 5 <= _text.Length &&
                string.Compare(_text, _position, "\\left", 0, 5, StringComparison.Ordinal) == 0;
        }

        private OfficeMathExpression ParseEnvironment() {
            string environment = ReadRequiredGroupName();
            if (environment != "bmatrix" && environment != "aligned" && environment != "gathered" && environment != "officeimostretchstack") {
                throw Error("Unsupported mathematical environment '" + environment + "'.");
            }
            int end = FindMatchingEnvironmentEnd(environment);
            if (end < 0) throw Error("The \\begin{" + environment + "} environment has no matching end.");
            string body = _text.Substring(_position, end - _position);
            _position = end + 4;
            string closing = ReadRequiredGroupName();
            if (!string.Equals(environment, closing, StringComparison.Ordinal)) throw Error("Mismatched mathematical environment end.");

            IReadOnlyList<IReadOnlyList<string>> sourceRows = SplitEnvironmentBody(body);
            int rows = Math.Max(1, sourceRows.Count);
            int columns = Math.Max(1, sourceRows.Count == 0 ? 1 : sourceRows.Max(row => row.Count));
            var cells = new List<OfficeMathExpression>(rows * columns);
            if (sourceRows.Count == 0) cells.Add(OfficeMath.Text(string.Empty));
            foreach (IReadOnlyList<string> row in sourceRows) {
                for (int column = 0; column < columns; column++) {
                    cells.Add(column < row.Count && !string.IsNullOrWhiteSpace(row[column])
                        ? new LatexParser(row[column], _maximumDepth, _depth).Parse()
                        : OfficeMath.Text(string.Empty));
                }
            }
            if (environment == "bmatrix") return OfficeMath.Matrix(rows, columns, cells.ToArray());
            if (environment == "aligned") return OfficeMath.EquationArray(rows, columns, cells.ToArray());
            OfficeMathExpression[] stackRows = Enumerable.Range(0, rows)
                .Select(row => CollapseRow(Enumerable.Range(0, columns).Select(column => cells[row * columns + column])))
                .ToArray();
            return environment == "officeimostretchstack" ? OfficeMath.StretchStack(stackRows) : OfficeMath.Stack(stackRows);
        }

        private string ReadRequiredGroupName() {
            SkipWhitespace();
            Require('{');
            int start = _position;
            while (_position < _text.Length && _text[_position] != '}') _position++;
            if (_position >= _text.Length) throw Error("An environment name is missing a closing brace.");
            string value = _text.Substring(start, _position - start);
            _position++;
            return value;
        }

        private int FindMatchingEnvironmentEnd(string environment) {
            var stack = new Stack<string>();
            stack.Push(environment);
            for (int index = _position; index < _text.Length; index++) {
                if (_text[index] != '\\' || !TryReadEnvironmentCommand(_text, index, out bool begin, out string? name, out int next)) continue;
                if (begin) stack.Push(name!);
                else {
                    if (stack.Count == 0 || !string.Equals(stack.Peek(), name, StringComparison.Ordinal)) return -1;
                    stack.Pop();
                    if (stack.Count == 0) return index;
                }
                index = next - 1;
            }
            return -1;
        }

        private static IReadOnlyList<IReadOnlyList<string>> SplitEnvironmentBody(string body) {
            var rows = new List<IReadOnlyList<string>>();
            var row = new List<string>();
            var cell = new StringBuilder();
            int braceDepth = 0;
            int environmentDepth = 0;
            for (int index = 0; index < body.Length; index++) {
                char value = body[index];
                if (value == '\\' && TryReadEnvironmentCommand(body, index, out bool begin, out _, out int next)) {
                    environmentDepth += begin ? 1 : -1;
                    cell.Append(body, index, next - index);
                    index = next - 1;
                    continue;
                }
                if (value == '\\' && index + 1 < body.Length) {
                    if (body[index + 1] == '\\' && braceDepth == 0 && environmentDepth == 0) {
                        row.Add(cell.ToString());
                        cell.Clear();
                        rows.Add(row.ToArray());
                        row = new List<string>();
                        index++;
                        continue;
                    }
                    cell.Append(value).Append(body[++index]);
                    continue;
                }
                if (value == '{') braceDepth++;
                else if (value == '}' && braceDepth > 0) braceDepth--;
                if (value == '&' && braceDepth == 0 && environmentDepth == 0) {
                    row.Add(cell.ToString());
                    cell.Clear();
                } else cell.Append(value);
            }
            row.Add(cell.ToString());
            rows.Add(row.ToArray());
            return rows;
        }

        private static bool TryReadEnvironmentCommand(string text, int index, out bool begin, out string? name, out int next) {
            begin = false;
            name = null;
            next = index;
            string prefix;
            if (index + 7 <= text.Length && string.Compare(text, index, "\\begin{", 0, 7, StringComparison.Ordinal) == 0) {
                begin = true;
                prefix = "\\begin{";
            } else if (index + 5 <= text.Length && string.Compare(text, index, "\\end{", 0, 5, StringComparison.Ordinal) == 0) {
                prefix = "\\end{";
            } else return false;
            int start = index + prefix.Length;
            int close = text.IndexOf('}', start);
            if (close < 0) return false;
            name = text.Substring(start, close - start);
            next = close + 1;
            return true;
        }

        private OfficeMathExpression ParseScript() {
            SkipWhitespace();
            if (_position >= _text.Length) throw Error("A script marker requires an operand.");
            return Peek('{') ? ParseRequiredGroup() : ParseAtom();
        }

        private OfficeMathExpression ParseRequiredGroup() {
            SkipWhitespace();
            if (!Peek('{')) throw Error("A braced operand is required.");
            _position++;
            OfficeMathExpression expression = ParseSequence('}');
            Require('}');
            return expression;
        }

        private OfficeMathExpression ParseTextGroup() {
            SkipWhitespace();
            Require('{');
            var builder = new StringBuilder();
            int nestedBraces = 0;
            while (_position < _text.Length) {
                char value = _text[_position++];
                if (value == '{') {
                    nestedBraces++;
                    builder.Append(value);
                    continue;
                }
                if (value == '}') {
                    if (nestedBraces == 0) return OfficeMath.Text(builder.ToString());
                    nestedBraces--;
                    builder.Append(value);
                    continue;
                }
                if (value != '\\') {
                    builder.Append(value);
                    continue;
                }

                string escaped = ReadCommandName();
                if (escaped == "backslash") {
                    builder.Append('\\');
                    if (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++;
                } else if (escaped == "_" || escaped == "^" || escaped == "{" || escaped == "}" ||
                    escaped == "#" || escaped == "%" || escaped == "&") {
                    builder.Append(escaped);
                } else {
                    builder.Append('\\').Append(escaped);
                }
            }
            throw Error("A text group is missing a closing brace.");
        }

        private string ReadCommandName() {
            if (_position >= _text.Length) throw Error("An incomplete command was found.");
            if (!char.IsLetter(_text[_position])) return _text[_position++].ToString(CultureInfo.InvariantCulture);
            int start = _position;
            while (_position < _text.Length && char.IsLetter(_text[_position])) _position++;
            return _text.Substring(start, _position - start);
        }

        private string ReadDelimiter() {
            SkipWhitespace();
            if (_position >= _text.Length) throw Error("A delimiter is required.");
            if (_text[_position] == '\\') {
                _position++;
                string command = ReadCommandName();
                return command == "{" ? "{" : command == "}" ? "}" : command;
            }
            return _text[_position++].ToString(CultureInfo.InvariantCulture);
        }

        private int FindMatchingRight() {
            int braceDepth = 0;
            int delimiterDepth = 0;
            for (int index = _position; index < _text.Length - 5; index++) {
                if (_text[index] == '\\') {
                    if (braceDepth == 0 && MatchesCommand(index, "left")) {
                        delimiterDepth++;
                        index += 4;
                        continue;
                    }
                    if (braceDepth == 0 && MatchesCommand(index, "right")) {
                        if (delimiterDepth == 0) return index;
                        delimiterDepth--;
                        index += 5;
                        continue;
                    }
                    // Escaped braces and command arguments must not affect structural brace depth.
                    if (index + 1 < _text.Length && (_text[index + 1] == '{' || _text[index + 1] == '}')) index++;
                    continue;
                }
                if (_text[index] == '{') braceDepth++;
                else if (_text[index] == '}' && braceDepth > 0) braceDepth--;
            }
            return -1;
        }

        private bool MatchesCommand(int index, string command) {
            int start = index + 1;
            if (start + command.Length > _text.Length ||
                string.Compare(_text, start, command, 0, command.Length, StringComparison.Ordinal) != 0) return false;
            int end = start + command.Length;
            return end >= _text.Length || !char.IsLetter(_text[end]);
        }

        private string ReadWhile(char first, Func<char, bool> predicate) {
            int start = _position - 1;
            while (_position < _text.Length && predicate(_text[_position])) _position++;
            return _text.Substring(start, _position - start);
        }

        private void SkipWhitespace() { while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++; }
        private bool Peek(char value) => _position < _text.Length && _text[_position] == value;
        private void Require(char value) { if (!Peek(value)) throw Error("Expected '" + value + "'."); _position++; }
        private FormatException Error(string message) => new FormatException(message + " Position " + _position.ToString(CultureInfo.InvariantCulture) + ".");
        private static bool IsNaryExpression(OfficeMathExpression expression) => expression.Kind == OfficeMathKind.Operator && IsNarySymbol(expression.Text);
    }
}
