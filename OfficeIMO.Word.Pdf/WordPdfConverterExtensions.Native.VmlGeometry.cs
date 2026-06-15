using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const string NativeVmlNumberPattern = @"[+-]?(?:\d+(?:\.\d*)?|\.\d+)";

        private static OpenXmlElement? GetNativeVmlReferencedShapeTypeElement(OpenXmlElement element) {
            string? type = GetNativeOpenXmlAttribute(element, "type");
            if (string.IsNullOrWhiteSpace(type) || !type!.TrimStart().StartsWith("#", StringComparison.Ordinal)) {
                return null;
            }

            string id = type.Trim().Substring(1);
            if (id.Length == 0) {
                return null;
            }

            OpenXmlElement root = element;
            while (root.Parent != null) {
                root = root.Parent;
            }

            return root.Descendants().FirstOrDefault(descendant =>
                descendant.NamespaceUri == "urn:schemas-microsoft-com:vml" &&
                descendant.LocalName.Equals("shapetype", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GetNativeOpenXmlAttribute(descendant, "id"), id, StringComparison.OrdinalIgnoreCase));
        }

        private static (double Width, double Height) GetNativeVmlCoordSize(OpenXmlElement element, OpenXmlElement? shapeType, double fallbackWidth, double fallbackHeight) {
            string? coordSize = GetNativeOpenXmlAttribute(element, "coordsize");
            if (string.IsNullOrWhiteSpace(coordSize) && shapeType != null) {
                coordSize = GetNativeOpenXmlAttribute(shapeType, "coordsize");
            }

            if (string.IsNullOrWhiteSpace(coordSize)) {
                return (fallbackWidth, fallbackHeight);
            }

            string[] parts = coordSize!.Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double width) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double height) &&
                width > 0D &&
                height > 0D) {
                return (width, height);
            }

            return (fallbackWidth, fallbackHeight);
        }

        private static IReadOnlyDictionary<string, double> GetNativeVmlFormulaValues(OpenXmlElement element, OpenXmlElement? shapeType, double coordWidth, double coordHeight) {
            var values = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            AddNativeVmlGeometryOperands(values, coordWidth, coordHeight);
            AddNativeVmlAdjustments(values, shapeType);
            AddNativeVmlAdjustments(values, element);

            OpenXmlElement formulaSource = shapeType ?? element;
            int index = 0;
            foreach (OpenXmlElement formula in formulaSource.Descendants().Where(descendant =>
                         descendant.NamespaceUri == "urn:schemas-microsoft-com:vml" &&
                         descendant.LocalName.Equals("f", StringComparison.OrdinalIgnoreCase))) {
                string? equation = GetNativeOpenXmlAttribute(formula, "eqn");
                if (string.IsNullOrWhiteSpace(equation)) {
                    continue;
                }

                if (TryEvaluateNativeVmlFormula(equation!, values, coordWidth, coordHeight, out double result)) {
                    values["@" + index.ToString(CultureInfo.InvariantCulture)] = result;
                }

                index++;
            }

            return values;
        }

        private static void AddNativeVmlGeometryOperands(IDictionary<string, double> values, double coordWidth, double coordHeight) {
            values["left"] = 0D;
            values["top"] = 0D;
            values["width"] = coordWidth;
            values["right"] = coordWidth;
            values["height"] = coordHeight;
            values["bottom"] = coordHeight;
            values["center"] = coordWidth / 2D;
            values["middle"] = coordHeight / 2D;
        }

        private static void AddNativeVmlAdjustments(IDictionary<string, double> values, OpenXmlElement? element) {
            if (element == null) {
                return;
            }

            string? value = GetNativeOpenXmlAttribute(element, "adj");
            if (string.IsNullOrWhiteSpace(value)) {
                value = GetNativeOpenXmlAttribute(element, "adjustment");
            }

            if (string.IsNullOrWhiteSpace(value)) {
                return;
            }

            string[] parts = value!.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++) {
                double? parsed = ParseNativeVmlDouble(parts[i]);
                if (parsed.HasValue) {
                    values["#" + i.ToString(CultureInfo.InvariantCulture)] = parsed.Value;
                }
            }
        }

        private static bool TryEvaluateNativeVmlFormula(string equation, IReadOnlyDictionary<string, double> values, double coordWidth, double coordHeight, out double result) {
            result = 0D;
            MatchCollection matches = Regex.Matches(equation, @"[^\s,]+");
            if (matches.Count == 0) {
                return false;
            }

            string op = matches[0].Value.ToLowerInvariant();
            double Read(int index, double fallback = 0D) =>
                index < matches.Count && TryResolveNativeVmlFormulaOperand(matches[index].Value, values, coordWidth, coordHeight, out double value)
                    ? value
                    : fallback;

            if (op == "val") {
                result = Read(1);
                return true;
            }

            if (op == "sum") {
                result = Read(1) + Read(2) - Read(3);
                return true;
            }

            if (op == "prod") {
                double divisor = Read(3, 1D);
                result = Math.Abs(divisor) > double.Epsilon ? Read(1) * Read(2, 1D) / divisor : 0D;
                return true;
            }

            if (op == "mid") {
                result = (Read(1) + Read(2)) / 2D;
                return true;
            }

            if (op == "abs") {
                result = Math.Abs(Read(1));
                return true;
            }

            if (op == "min") {
                result = Math.Min(Read(1), Read(2));
                return true;
            }

            if (op == "max") {
                result = Math.Max(Read(1), Read(2));
                return true;
            }

            if (op == "if") {
                result = Read(1) > 0D ? Read(2) : Read(3);
                return true;
            }

            if (op == "sqrt") {
                result = Math.Sqrt(Math.Max(0D, Read(1)));
                return true;
            }

            if (op == "mod") {
                double x = Read(1);
                double y = Read(2);
                double z = Read(3);
                result = Math.Sqrt(x * x + y * y + z * z);
                return true;
            }

            return false;
        }

        private static bool TryResolveNativeVmlFormulaOperand(string value, IReadOnlyDictionary<string, double> values, double coordWidth, double coordHeight, out double result) {
            if (values.TryGetValue(value, out result)) {
                return true;
            }

            if (value.Equals("width", StringComparison.OrdinalIgnoreCase)) {
                result = coordWidth;
                return true;
            }

            if (value.Equals("right", StringComparison.OrdinalIgnoreCase)) {
                result = coordWidth;
                return true;
            }

            if (value.Equals("height", StringComparison.OrdinalIgnoreCase)) {
                result = coordHeight;
                return true;
            }

            if (value.Equals("bottom", StringComparison.OrdinalIgnoreCase)) {
                result = coordHeight;
                return true;
            }

            if (value.Equals("left", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("top", StringComparison.OrdinalIgnoreCase)) {
                result = 0D;
                return true;
            }

            if (value.Equals("center", StringComparison.OrdinalIgnoreCase)) {
                result = coordWidth / 2D;
                return true;
            }

            if (value.Equals("middle", StringComparison.OrdinalIgnoreCase)) {
                result = coordHeight / 2D;
                return true;
            }

            double? parsed = ParseNativeVmlDouble(value);
            if (parsed.HasValue) {
                result = parsed.Value;
                return true;
            }

            result = 0D;
            return false;
        }

        private static IReadOnlyList<string> TokenizeNativeVmlPathParts(string segment, IReadOnlyDictionary<string, double> formulaValues) {
            var parts = new List<string>();
            int index = 0;
            bool hasTokenSinceComma = false;
            bool endedWithComma = false;

            while (index < segment.Length) {
                char current = segment[index];
                if (current == ',') {
                    if (!hasTokenSinceComma) {
                        parts.Add("0");
                    }

                    hasTokenSinceComma = false;
                    endedWithComma = true;
                    index++;
                    continue;
                }

                if (char.IsWhiteSpace(current)) {
                    index++;
                    continue;
                }

                string token = ReadNativeVmlPathPart(segment, ref index, formulaValues);
                if (token.Length > 0) {
                    parts.Add(token);
                    hasTokenSinceComma = true;
                    endedWithComma = false;
                }
            }

            if (endedWithComma) {
                parts.Add("0");
            }

            return parts;
        }

        private static int FindNativeVmlPathSegmentEnd(string path, int index, IReadOnlyDictionary<string, double> formulaValues) {
            while (index < path.Length) {
                char current = path[index];
                if (char.IsWhiteSpace(current) || current == ',') {
                    index++;
                    continue;
                }

                if (current == '@' || current == '#') {
                    index++;
                    while (index < path.Length && char.IsDigit(path[index])) {
                        index++;
                    }

                    continue;
                }

                if (IsNativeVmlPathNumberStart(current)) {
                    index = SkipNativeVmlPathNumber(path, index);
                    continue;
                }

                if (char.IsLetter(current)) {
                    int tokenEnd = index + 1;
                    while (tokenEnd < path.Length && char.IsLetter(path[tokenEnd])) {
                        tokenEnd++;
                    }

                    string token = path.Substring(index, tokenEnd - index);
                    if (formulaValues.ContainsKey(token)) {
                        index = tokenEnd;
                        continue;
                    }

                    if (IsNativeVmlPathCommand(char.ToLowerInvariant(current))) {
                        return index;
                    }

                    index++;
                    continue;
                }

                index++;
            }

            return index;
        }

        private static string ReadNativeVmlPathPart(string segment, ref int index, IReadOnlyDictionary<string, double> formulaValues) {
            int start = index;
            if (segment[index] == '@' || segment[index] == '#') {
                index++;
                while (index < segment.Length && char.IsDigit(segment[index])) {
                    index++;
                }

                return segment.Substring(start, index - start);
            }

            if (char.IsLetter(segment[index])) {
                index++;
                while (index < segment.Length && char.IsLetter(segment[index])) {
                    index++;
                }

                string token = segment.Substring(start, index - start);
                return formulaValues.ContainsKey(token) ? token : string.Empty;
            }

            if (!IsNativeVmlPathNumberStart(segment[index])) {
                index++;
                return string.Empty;
            }

            if (segment[index] == '+' || segment[index] == '-') {
                index++;
            }

            while (index < segment.Length) {
                char current = segment[index];
                if (current == ',' || char.IsWhiteSpace(current) || current == '@' || current == '#') {
                    break;
                }

                index++;
            }

            return segment.Substring(start, index - start);
        }

        private static int SkipNativeVmlPathNumber(string value, int index) {
            if (value[index] == '+' || value[index] == '-') {
                index++;
            }

            while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.')) {
                index++;
            }

            if (index < value.Length && (value[index] == 'e' || value[index] == 'E')) {
                int exponentStart = index;
                int exponentIndex = index + 1;
                if (exponentIndex < value.Length && (value[exponentIndex] == '+' || value[exponentIndex] == '-')) {
                    exponentIndex++;
                }

                int digitsStart = exponentIndex;
                while (exponentIndex < value.Length && char.IsDigit(value[exponentIndex])) {
                    exponentIndex++;
                }

                if (exponentIndex > digitsStart) {
                    index = exponentIndex;
                } else {
                    index = exponentStart;
                }
            }

            return index;
        }

        private static bool IsNativeVmlPathNumberStart(char value) =>
            char.IsDigit(value) || value == '.' || value == '+' || value == '-';

        private static string ResolveNativeVmlFormulaReferences(string value, IReadOnlyDictionary<string, double> formulaValues) =>
            Regex.Replace(value, @"[@#]\d+", match =>
                formulaValues.TryGetValue(match.Value, out double resolved)
                    ? resolved.ToString(CultureInfo.InvariantCulture)
                    : "0");
    }
}
