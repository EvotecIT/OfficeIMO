using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioShapeGeometry {

        private static bool TryReadCell(XElement row, XNamespace ns, string name, VisioShape shape, out double value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (cell == null) {
                value = 0D;
                return false;
            }

            if (TryParseCellLiteral(cell.Attribute("V")?.Value, shape, out value)) {
                return true;
            }

            if (TryParseCellLiteral(cell.Attribute("F")?.Value, shape, out value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryReadRawCell(XElement row, XNamespace ns, string name, out double value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (cell == null) {
                value = 0D;
                return false;
            }

            if (TryParseLiteralWithoutShape(cell.Attribute("V")?.Value, out value)) {
                return true;
            }

            return TryParseLiteralWithoutShape(cell.Attribute("F")?.Value, out value);
        }

        private static bool TryReadFormulaCell(XElement row, XNamespace ns, string name, out string? value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            value = cell?.Attribute("F")?.Value ?? cell?.Attribute("V")?.Value;
            return !string.IsNullOrWhiteSpace(value);
        }

        private static bool TryReadBooleanCell(XElement row, XNamespace ns, string name, VisioShape shape, out bool value) {
            if (TryReadCell(row, ns, name, shape, out double numeric)) {
                value = Math.Abs(numeric) > 1e-9;
                return true;
            }

            value = false;
            return false;
        }

        private static bool TryParseCellLiteral(string? raw, VisioShape shape, out double value) {
            raw = NormalizeCellLiteral(raw);
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                IsFinite(value)) {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(raw) &&
                CellExpressionParser.TryEvaluate(raw!, shape.Width, shape.Height, shape.LocPinX, shape.LocPinY, shape.PinX, shape.PinY, shape.Angle, out value) &&
                IsFinite(value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryParseLiteralWithoutShape(string? raw, out double value) {
            raw = NormalizeCellLiteral(raw);
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                IsFinite(value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryParsePolylineFormula(string? raw, VisioShape shape, out List<(double X, double Y)> points) {
            points = new List<(double X, double Y)>();
            if (!TryParseFunctionArguments(raw, "POLYLINE", out List<string> parts)) {
                return false;
            }

            if (parts.Count < 4 || parts.Count % 2 != 0) {
                return false;
            }

            if (!TryParsePolylineArgument(parts[0], shape, out double rawXType) ||
                !TryParsePolylineArgument(parts[1], shape, out double rawYType)) {
                return false;
            }

            bool xIsLocal = Math.Abs(rawXType) > 1e-9;
            bool yIsLocal = Math.Abs(rawYType) > 1e-9;
            for (int i = 2; i < parts.Count; i += 2) {
                if (!TryParsePolylineArgument(parts[i], shape, out double rawX) ||
                    !TryParsePolylineArgument(parts[i + 1], shape, out double rawY)) {
                    return false;
                }

                points.Add((xIsLocal ? rawX : rawX * shape.Width, yIsLocal ? rawY : rawY * shape.Height));
            }

            return points.Count > 0;
        }

        private static bool TryParsePolylineArgument(string raw, VisioShape shape, out double value) =>
            TryParseCellLiteral(raw, shape, out value);

        private static bool TryParseNurbsFormula(
            string? raw,
            VisioShape shape,
            (double X, double Y) start,
            (double X, double Y) end,
            double firstKnot,
            double firstWeight,
            double secondLastKnot,
            double lastWeight,
            out NurbsCurve? curve) {
            curve = null;
            if (!TryParseFunctionArguments(raw, "NURBS", out List<string> arguments) ||
                arguments.Count < 8 ||
                (arguments.Count - 4) % 4 != 0 ||
                !TryParsePolylineArgument(arguments[0], shape, out double lastKnot) ||
                !TryParsePolylineArgument(arguments[1], shape, out double rawDegree) ||
                !TryParsePolylineArgument(arguments[2], shape, out double rawXType) ||
                !TryParsePolylineArgument(arguments[3], shape, out double rawYType)) {
                return false;
            }

            int degree = (int)Math.Round(rawDegree);
            if (degree < 1 || degree > 25) {
                return false;
            }

            bool xIsLocal = Math.Abs(rawXType) > 1e-9;
            bool yIsLocal = Math.Abs(rawYType) > 1e-9;
            List<(double X, double Y)> controlPoints = new() { start };
            List<double> weights = new() { firstWeight };
            List<double> suppliedKnots = new() { firstKnot };
            for (int i = 4; i < arguments.Count; i += 4) {
                if (!TryParsePolylineArgument(arguments[i], shape, out double rawX) ||
                    !TryParsePolylineArgument(arguments[i + 1], shape, out double rawY) ||
                    !TryParsePolylineArgument(arguments[i + 2], shape, out double knot) ||
                    !TryParsePolylineArgument(arguments[i + 3], shape, out double weight)) {
                    return false;
                }

                controlPoints.Add((xIsLocal ? rawX : rawX * shape.Width, yIsLocal ? rawY : rawY * shape.Height));
                weights.Add(weight);
                suppliedKnots.Add(knot);
            }

            controlPoints.Add(end);
            weights.Add(lastWeight);
            suppliedKnots.Add(secondLastKnot);
            suppliedKnots.Add(lastKnot);
            if (controlPoints.Count <= degree) {
                degree = controlPoints.Count - 1;
            }

            if (degree < 1 || weights.Any(weight => !IsFinite(weight) || weight <= 0D)) {
                return false;
            }

            List<double> knots = TryUseSuppliedKnotVector(suppliedKnots, controlPoints.Count, degree, out List<double>? normalizedKnots)
                ? normalizedKnots!
                : BuildClampedKnotVector(controlPoints.Count, degree, firstKnot, lastKnot);
            curve = new NurbsCurve(controlPoints, weights, knots, degree);
            return true;
        }

        private static bool TryParseFunctionArguments(string? raw, string functionName, out List<string> arguments) {
            arguments = new List<string>();
            raw = NormalizeCellLiteral(raw);
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            string formula = raw!.Trim();
            if (formula.StartsWith("GUARD(", StringComparison.OrdinalIgnoreCase) && formula.EndsWith(")", StringComparison.Ordinal)) {
                formula = formula.Substring(6, formula.Length - 7).Trim();
            }

            if (!formula.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int openIndex = formula.IndexOf('(');
            int closeIndex = formula.LastIndexOf(')');
            if (openIndex < 0 || closeIndex <= openIndex) {
                return false;
            }

            string argumentsText = formula.Substring(openIndex + 1, closeIndex - openIndex - 1);
            int depth = 0;
            int start = 0;
            for (int i = 0; i < argumentsText.Length; i++) {
                char current = argumentsText[i];
                if (current == '(') {
                    depth++;
                } else if (current == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }
                } else if (current == ',' && depth == 0) {
                    arguments.Add(argumentsText.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }

            arguments.Add(argumentsText.Substring(start).Trim());
            return depth == 0 && arguments.All(argument => argument.Length > 0);
        }

        private static bool TryUseSuppliedKnotVector(List<double> suppliedKnots, int controlPointCount, int degree, out List<double>? knots) {
            knots = null;
            if (suppliedKnots.Any(knot => !IsFinite(knot))) {
                return false;
            }

            for (int i = 1; i < suppliedKnots.Count; i++) {
                if (suppliedKnots[i] + 1e-9 < suppliedKnots[i - 1]) {
                    return false;
                }
            }

            int expectedFullCount = controlPointCount + degree + 1;
            if (suppliedKnots.Count == expectedFullCount) {
                knots = new List<double>(suppliedKnots);
                return IsUsableKnotDomain(knots, controlPointCount, degree);
            }

            if (suppliedKnots.Count == controlPointCount + 1) {
                List<double> expanded = new(suppliedKnots);
                double last = suppliedKnots[suppliedKnots.Count - 1];
                for (int i = 0; i < degree; i++) {
                    expanded.Add(last);
                }

                if (expanded.Count == expectedFullCount && IsUsableKnotDomain(expanded, controlPointCount, degree)) {
                    knots = expanded;
                    return true;
                }
            }

            return false;
        }

        private static bool IsUsableKnotDomain(List<double> knots, int controlPointCount, int degree) {
            if (knots.Count <= controlPointCount || knots.Count <= degree) {
                return false;
            }

            double start = knots[degree];
            double end = knots[controlPointCount];
            return IsFinite(start) && IsFinite(end) && end > start;
        }

        private static List<double> BuildClampedKnotVector(int controlPointCount, int degree, double firstKnot, double lastKnot) {
            double start = IsFinite(firstKnot) ? firstKnot : 0D;
            double end = IsFinite(lastKnot) && lastKnot > start ? lastKnot : start + 1D;
            int knotCount = controlPointCount + degree + 1;
            List<double> knots = new(knotCount);
            for (int i = 0; i < knotCount; i++) {
                if (i <= degree) {
                    knots.Add(start);
                } else if (i >= controlPointCount) {
                    knots.Add(end);
                } else {
                    double step = (i - degree) / (double)(controlPointCount - degree);
                    knots.Add(start + ((end - start) * step));
                }
            }

            return knots;
        }

        private static string? NormalizeCellLiteral(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return value;
            }

            string trimmed = value!.Trim();
            string normalized = trimmed.StartsWith("=", StringComparison.Ordinal)
                ? trimmed.Substring(1)
                : trimmed;
            return StripOuterGuard(normalized);
        }

        private static string StripOuterGuard(string value) {
            string normalized = value.Trim();
            while (TryGetFullFunctionArgument(normalized, "GUARD", out string? argument)) {
                normalized = argument!.Trim();
            }

            return normalized;
        }

        private static bool TryGetFullFunctionArgument(string value, string functionName, out string? argument) {
            argument = null;
            if (!value.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int openIndex = functionName.Length;
            while (openIndex < value.Length && char.IsWhiteSpace(value[openIndex])) {
                openIndex++;
            }

            if (openIndex >= value.Length || value[openIndex] != '(') {
                return false;
            }

            int depth = 0;
            for (int i = openIndex; i < value.Length; i++) {
                if (value[i] == '(') {
                    depth++;
                } else if (value[i] == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }

                    if (depth == 0) {
                        if (i != value.Length - 1) {
                            return false;
                        }

                        argument = value.Substring(openIndex + 1, i - openIndex - 1);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool NearlyEqual((double X, double Y) a, (double X, double Y) b) =>
            Math.Abs(a.X - b.X) <= 1e-9 &&
            Math.Abs(a.Y - b.Y) <= 1e-9;

        private static bool IsFinite(double value) =>
            !double.IsNaN(value) &&
            !double.IsInfinity(value);

        private sealed class NurbsCurve {
            internal NurbsCurve(List<(double X, double Y)> controlPoints, List<double> weights, List<double> knots, int degree) {
                ControlPoints = controlPoints;
                Weights = weights;
                Knots = knots;
                Degree = degree;
            }

            internal List<(double X, double Y)> ControlPoints { get; }

            internal List<double> Weights { get; }

            internal List<double> Knots { get; }

            internal int Degree { get; }
        }
    }
}
