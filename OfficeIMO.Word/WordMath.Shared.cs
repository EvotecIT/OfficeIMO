using OfficeIMO.Drawing;
using System.Globalization;
using System.Xml.Linq;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Word {
    internal static partial class WordMath {
        private static readonly XNamespace OmmlNamespace = MathNamespace;

        internal static OfficeMathExpression ToExpression(OpenXmlElement element) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            if (element is M.Text text) return ClassifyToken(text.Text ?? string.Empty);

            switch (element.LocalName) {
                case "f":
                    OfficeMathExpression numerator = ExpressionFromChild(element, "num");
                    OfficeMathExpression denominator = ExpressionFromChild(element, "den");
                    return IsSkewedFraction(element)
                        ? OfficeIMO.Drawing.OfficeMath.SlashedFraction(numerator, denominator)
                        : OfficeIMO.Drawing.OfficeMath.Fraction(numerator, denominator);
                case "rad":
                    OfficeMathExpression radicand = ExpressionFromChild(element, "e");
                    OfficeMathExpression degree = ExpressionFromChild(element, "deg");
                    return degree.ToPlainText().Length == 0
                        ? OfficeIMO.Drawing.OfficeMath.Radical(radicand)
                        : OfficeIMO.Drawing.OfficeMath.Radical(radicand, degree);
                case "sSup":
                    return OfficeIMO.Drawing.OfficeMath.Superscript(ExpressionFromChild(element, "e"), ExpressionFromChild(element, "sup"));
                case "sSub":
                    return OfficeIMO.Drawing.OfficeMath.Subscript(ExpressionFromChild(element, "e"), ExpressionFromChild(element, "sub"));
                case "sSubSup":
                    return OfficeIMO.Drawing.OfficeMath.SubSuperscript(
                        ExpressionFromChild(element, "e"),
                        ExpressionFromChild(element, "sub"),
                        ExpressionFromChild(element, "sup"));
                case "sPre":
                    return OfficeIMO.Drawing.OfficeMath.LeftSubSuperscript(
                        ExpressionFromChild(element, "e"),
                        ExpressionFromChild(element, "sub"),
                        ExpressionFromChild(element, "sup"));
                case "nary":
                case "int":
                    return NaryExpression(element);
                case "func":
                    string functionName = ReadChildText(element, "fName");
                    return string.IsNullOrWhiteSpace(functionName)
                        ? ExpressionFromChild(element, "e")
                        : OfficeIMO.Drawing.OfficeMath.Function(functionName, ExpressionFromChild(element, "e"));
                case "acc":
                    return OfficeIMO.Drawing.OfficeMath.Accent(ExpressionFromChild(element, "e"), ReadCharacterOrDefault(element, "chr", "\u0302"));
                case "bar":
                    return string.Equals(ReadCharacter(element, "pos").Value, "bot", StringComparison.OrdinalIgnoreCase)
                        ? OfficeIMO.Drawing.OfficeMath.Underbar(ExpressionFromChild(element, "e"))
                        : OfficeIMO.Drawing.OfficeMath.Overbar(ExpressionFromChild(element, "e"));
                case "d":
                    return DelimitedExpression(element);
                case "groupChr":
                    return OfficeIMO.Drawing.OfficeMath.Accent(ExpressionFromChild(element, "e"), ReadCharacterOrDefault(element, "chr", "\u23df"));
                case "m":
                    return MatrixExpression(element);
                case "eqArr":
                    return EquationArrayExpression(element);
                case "limLow":
                    return OfficeIMO.Drawing.OfficeMath.LowerLimit(ExpressionFromChild(element, "e"), ExpressionFromChild(element, "lim"));
                case "limUpp":
                    return OfficeIMO.Drawing.OfficeMath.UpperLimit(ExpressionFromChild(element, "e"), ExpressionFromChild(element, "lim"));
                case "borderBox":
                    return OfficeIMO.Drawing.OfficeMath.Box(ExpressionFromChild(element, "e"));
                case "phant":
                    return OfficeIMO.Drawing.OfficeMath.Phantom(ExpressionFromChild(element, "e"));
                default:
                    return CollapseExpressions(element.ChildElements
                        .Where(child => !child.LocalName.EndsWith("Pr", StringComparison.Ordinal))
                        .Select(ToExpression));
            }
        }

        internal static string ToOmml(OfficeMathExpression expression, bool display) {
            if (expression == null) throw new ArgumentNullException(nameof(expression));
            var officeMath = new XElement(OmmlNamespace + "oMath", new XAttribute(XNamespace.Xmlns + "m", OmmlNamespace));
            AppendOmml(officeMath, expression);
            if (!display) return officeMath.ToString(SaveOptions.DisableFormatting);
            var paragraph = new XElement(OmmlNamespace + "oMathPara", new XAttribute(XNamespace.Xmlns + "m", OmmlNamespace), officeMath);
            return paragraph.ToString(SaveOptions.DisableFormatting);
        }

        private static void AppendOmml(XElement parent, OfficeMathExpression expression) {
            switch (expression.Kind) {
                case OfficeMathKind.Text:
                case OfficeMathKind.Identifier:
                case OfficeMathKind.Number:
                case OfficeMathKind.Operator:
                    parent.Add(MathRun(expression.Text ?? string.Empty));
                    return;
                case OfficeMathKind.Row:
                    foreach (OfficeMathExpression child in expression.Children) AppendOmml(parent, child);
                    return;
                case OfficeMathKind.Fraction:
                    parent.Add(Composite("f", Container("num", expression.Children[0]), Container("den", expression.Children[1])));
                    return;
                case OfficeMathKind.SlashedFraction:
                    parent.Add(Composite("f", CharacterProperties("fPr", "type", "skw"),
                        Container("num", expression.Children[0]), Container("den", expression.Children[1])));
                    return;
                case OfficeMathKind.Radical:
                    XElement degree = new XElement(OmmlNamespace + "deg");
                    if (expression.Children.Count == 2) AppendOmml(degree, expression.Children[1]);
                    parent.Add(Composite("rad", degree, Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Superscript:
                    parent.Add(Composite("sSup", Container("e", expression.Children[0]), Container("sup", expression.Children[1])));
                    return;
                case OfficeMathKind.Subscript:
                    parent.Add(Composite("sSub", Container("e", expression.Children[0]), Container("sub", expression.Children[1])));
                    return;
                case OfficeMathKind.SubSuperscript:
                    parent.Add(Composite("sSubSup", Container("e", expression.Children[0]), Container("sub", expression.Children[1]), Container("sup", expression.Children[2])));
                    return;
                case OfficeMathKind.LeftSubSuperscript:
                    parent.Add(Composite("sPre", Container("sub", expression.Children[1]), Container("sup", expression.Children[2]), Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.LowerLimit:
                    parent.Add(Composite("limLow", Container("e", expression.Children[0]), Container("lim", expression.Children[1])));
                    return;
                case OfficeMathKind.UpperLimit:
                    parent.Add(Composite("limUpp", Container("e", expression.Children[0]), Container("lim", expression.Children[1])));
                    return;
                case OfficeMathKind.Nary:
                    AppendNaryOmml(parent, expression);
                    return;
                case OfficeMathKind.Delimited:
                    AppendDelimitedOmml(parent, expression);
                    return;
                case OfficeMathKind.DelimiterList:
                    AppendDelimiterListOmml(parent, expression);
                    return;
                case OfficeMathKind.Function:
                    parent.Add(Composite("func", Container("fName", OfficeIMO.Drawing.OfficeMath.Identifier(expression.Text ?? string.Empty)), Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Matrix:
                    AppendMatrixOmml(parent, expression);
                    return;
                case OfficeMathKind.EquationArray:
                    AppendEquationArrayOmml(parent, expression);
                    return;
                case OfficeMathKind.Accent:
                    parent.Add(Composite("acc", CharacterProperties("accPr", "chr", expression.Character ?? "\u0302"), Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Overbar:
                case OfficeMathKind.Underbar:
                    parent.Add(Composite("bar", CharacterProperties("barPr", "pos", expression.Kind == OfficeMathKind.Underbar ? "bot" : "top"), Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Box:
                    parent.Add(Composite("borderBox", Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Phantom:
                    parent.Add(Composite("phant", Container("e", expression.Children[0])));
                    return;
                case OfficeMathKind.Stack:
                case OfficeMathKind.StretchStack:
                    throw new NotSupportedException(
                        "Word OMML has no lossless representation for shared math node " + expression.Kind + ". " +
                        "Use an equation array explicitly when that lossy projection is intended.");
                default:
                    throw new NotSupportedException("Unsupported shared math node for OMML: " + expression.Kind + ".");
            }
        }

        private static void AppendNaryOmml(XElement parent, OfficeMathExpression expression) {
            var properties = CharacterProperties("naryPr", "chr", expression.Character ?? "∑");
            var nary = Composite("nary", properties);
            var subscript = new XElement(OmmlNamespace + "sub");
            var superscript = new XElement(OmmlNamespace + "sup");
            if (expression.Children.Count > 1) AppendOmml(subscript, expression.Children[1]);
            if (expression.Children.Count > 2) AppendOmml(superscript, expression.Children[2]);
            nary.Add(subscript, superscript, Container("e", expression.Children[0]));
            parent.Add(nary);
        }

        private static void AppendDelimitedOmml(XElement parent, OfficeMathExpression expression) {
            var properties = new XElement(OmmlNamespace + "dPr",
                CharacterValue("begChr", expression.Character ?? "("),
                CharacterValue("endChr", expression.SecondaryCharacter ?? ")"));
            parent.Add(Composite("d", properties, Container("e", expression.Children[0])));
        }

        private static void AppendDelimiterListOmml(XElement parent, OfficeMathExpression expression) {
            var properties = new XElement(OmmlNamespace + "dPr",
                CharacterValue("begChr", expression.Character ?? "("),
                CharacterValue("sepChr", expression.SeparatorCharacter ?? ","),
                CharacterValue("endChr", expression.SecondaryCharacter ?? ")"));
            var delimiter = Composite("d", properties);
            foreach (OfficeMathExpression child in expression.Children) delimiter.Add(Container("e", child));
            parent.Add(delimiter);
        }

        private static void AppendMatrixOmml(XElement parent, OfficeMathExpression expression) {
            var matrix = new XElement(OmmlNamespace + "m");
            for (int row = 0; row < expression.RowCount; row++) {
                var matrixRow = new XElement(OmmlNamespace + "mr");
                for (int column = 0; column < expression.ColumnCount; column++) {
                    matrixRow.Add(Container("e", expression.Children[row * expression.ColumnCount + column]));
                }
                matrix.Add(matrixRow);
            }
            parent.Add(matrix);
        }

        private static void AppendEquationArrayOmml(XElement parent, OfficeMathExpression expression) {
            var array = new XElement(OmmlNamespace + "eqArr");
            for (int row = 0; row < expression.RowCount; row++) {
                var rowExpression = new XElement(OmmlNamespace + "e");
                for (int column = 0; column < expression.ColumnCount; column++) {
                    if (column > 0) rowExpression.Add(AlignmentRun());
                    AppendOmml(rowExpression, expression.Children[row * expression.ColumnCount + column]);
                }
                array.Add(rowExpression);
            }
            parent.Add(array);
        }

        private static OfficeMathExpression NaryExpression(OpenXmlElement element) {
            string character = ReadCharacter(element, "chr").Present
                ? ReadCharacter(element, "chr").Value
                : element.LocalName == "int" ? "∫" : "∑";
            OfficeMathExpression content = ExpressionFromChild(element, "e");
            OfficeMathExpression lower = ExpressionFromChild(element, "sub");
            OfficeMathExpression upper = ExpressionFromChild(element, "sup");
            return OfficeIMO.Drawing.OfficeMath.Nary(
                character,
                content,
                lower.ToPlainText().Length == 0 ? null : lower,
                upper.ToPlainText().Length == 0 ? null : upper);
        }

        private static OfficeMathExpression DelimitedExpression(OpenXmlElement element) {
            MathCharacter begin = ReadCharacter(element, "begChr");
            MathCharacter end = ReadCharacter(element, "endChr");
            string separator = ReadDelimiterSeparator(element);
            OfficeMathExpression[] content = FindChildren(element, "e").Select(ToExpression).ToArray();
            if (content.Length > 1) {
                return OfficeIMO.Drawing.OfficeMath.DelimiterList(
                    begin.Present ? begin.Value : "(",
                    end.Present ? end.Value : ")",
                    separator,
                    content);
            }
            return OfficeIMO.Drawing.OfficeMath.Delimited(
                content.Length == 0 ? OfficeIMO.Drawing.OfficeMath.Text(string.Empty) : content[0],
                begin.Present ? begin.Value : "(",
                end.Present ? end.Value : ")");
        }

        private static bool IsSkewedFraction(OpenXmlElement element) {
            return ReadFractionType(element) == MathFractionType.Skewed;
        }

        private static OfficeMathExpression MatrixExpression(OpenXmlElement element) {
            List<OpenXmlElement> rows = FindChildren(element, "mr").ToList();
            int columns = rows.Count == 0 ? 1 : Math.Max(1, rows.Max(row => FindChildren(row, "e").Count()));
            var cells = new List<OfficeMathExpression>(Math.Max(1, rows.Count) * columns);
            if (rows.Count == 0) cells.Add(OfficeIMO.Drawing.OfficeMath.Text(string.Empty));
            foreach (OpenXmlElement row in rows) {
                List<OpenXmlElement> rowCells = FindChildren(row, "e").ToList();
                for (int column = 0; column < columns; column++) {
                    cells.Add(column < rowCells.Count ? ToExpression(rowCells[column]) : OfficeIMO.Drawing.OfficeMath.Text(string.Empty));
                }
            }
            return OfficeIMO.Drawing.OfficeMath.Matrix(Math.Max(1, rows.Count), columns, cells.ToArray());
        }

        private static OfficeMathExpression EquationArrayExpression(OpenXmlElement element) {
            List<OpenXmlElement> rows = FindChildren(element, "e").ToList();
            if (rows.Count == 0) return OfficeIMO.Drawing.OfficeMath.EquationArray(1, 1, OfficeIMO.Drawing.OfficeMath.Text(string.Empty));
            var splitRows = new List<IReadOnlyList<OfficeMathExpression>>(rows.Count);
            int columns = 1;
            foreach (OpenXmlElement row in rows) {
                var cells = new List<OfficeMathExpression>();
                var current = new List<OfficeMathExpression>();
                foreach (OpenXmlElement child in row.ChildElements) {
                    if (child.LocalName == "aln" || child.Descendants().Any(descendant =>
                        descendant.NamespaceUri == MathNamespace && descendant.LocalName == "aln")) {
                        cells.Add(CollapseExpressions(current));
                        current.Clear();
                    } else if (!child.LocalName.EndsWith("Pr", StringComparison.Ordinal)) current.Add(ToExpression(child));
                }
                cells.Add(CollapseExpressions(current));
                columns = Math.Max(columns, cells.Count);
                splitRows.Add(cells);
            }
            var flattened = new List<OfficeMathExpression>(rows.Count * columns);
            foreach (IReadOnlyList<OfficeMathExpression> row in splitRows) {
                for (int column = 0; column < columns; column++) {
                    flattened.Add(column < row.Count ? row[column] : OfficeIMO.Drawing.OfficeMath.Text(string.Empty));
                }
            }
            return OfficeIMO.Drawing.OfficeMath.EquationArray(rows.Count, columns, flattened.ToArray());
        }

        private static OfficeMathExpression ExpressionFromChild(OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstChild(element, localName);
            return child == null ? OfficeIMO.Drawing.OfficeMath.Text(string.Empty) : ToExpression(child);
        }

        private static OfficeMathExpression CollapseExpressions(IEnumerable<OfficeMathExpression> expressions) {
            OfficeMathExpression[] values = expressions.ToArray();
            if (values.Length == 0) return OfficeIMO.Drawing.OfficeMath.Text(string.Empty);
            return values.Length == 1 ? values[0] : OfficeIMO.Drawing.OfficeMath.Row(values);
        }

        private static OfficeMathExpression ClassifyToken(string value) {
            if (value.Length > 0 && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) return OfficeIMO.Drawing.OfficeMath.Number(value);
            if (value.Length > 0 && value.All(character => char.IsLetter(character))) return OfficeIMO.Drawing.OfficeMath.Identifier(value);
            if (value.Length > 0 && value.All(character => "+-−=*/×÷±∑∏∫≤≥≠→∞,.;:|".IndexOf(character) >= 0)) return OfficeIMO.Drawing.OfficeMath.Operator(value);
            return OfficeIMO.Drawing.OfficeMath.Text(value);
        }

        private static XElement Container(string name, OfficeMathExpression expression) {
            var container = new XElement(OmmlNamespace + name);
            AppendOmml(container, expression);
            return container;
        }

        private static XElement Composite(string name, params object[] content) => new XElement(OmmlNamespace + name, content);

        private static XElement CharacterProperties(string propertyName, string characterName, string value) =>
            new XElement(OmmlNamespace + propertyName, CharacterValue(characterName, value));

        private static XElement CharacterValue(string name, string value) =>
            new XElement(OmmlNamespace + name, new XAttribute(OmmlNamespace + "val", value));

        private static XElement MathRun(string value) {
            var text = new XElement(OmmlNamespace + "t", value);
            if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]))) {
                text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
            return new XElement(OmmlNamespace + "r", text);
        }

        private static XElement AlignmentRun() =>
            new XElement(OmmlNamespace + "r",
                new XElement(OmmlNamespace + "rPr", new XElement(OmmlNamespace + "aln")),
                new XElement(OmmlNamespace + "t", string.Empty));
    }
}
