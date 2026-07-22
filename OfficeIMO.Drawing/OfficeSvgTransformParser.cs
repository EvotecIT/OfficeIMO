using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

internal static class OfficeSvgTransformParser {
    private const int MaximumTransformOperations = 256;

    internal static bool TryParse(string? value, out OfficeTransform transform) {
        transform = OfficeTransform.Identity;
        if (string.IsNullOrWhiteSpace(value)) return true;
        int index = 0;
        int operationCount = 0;
        while (SkipSeparators(value!, ref index)) {
            if (++operationCount > MaximumTransformOperations) return false;
            int nameStart = index;
            while (index < value!.Length && char.IsLetter(value[index])) index++;
            if (index == nameStart) return false;
            string name = value.Substring(nameStart, index - nameStart).ToLowerInvariant();
            SkipWhitespace(value, ref index);
            if (index >= value.Length || value[index++] != '(') return false;
            int argumentsStart = index;
            int close = value.IndexOf(')', index);
            if (close < 0) return false;
            if (!TryParseNumbers(value.Substring(argumentsStart, close - argumentsStart), out IReadOnlyList<double> arguments)
                || !TryCreateTransform(name, arguments, out OfficeTransform current)) return false;
            transform = current.Then(transform);
            index = close + 1;
        }
        return true;
    }

    private static bool TryCreateTransform(string name, IReadOnlyList<double> values, out OfficeTransform transform) {
        transform = OfficeTransform.Identity;
        switch (name) {
            case "matrix":
                if (values.Count != 6) return false;
                transform = new OfficeTransform(values[0], values[1], values[2], values[3], values[4], values[5]);
                return true;
            case "translate":
                if (values.Count is < 1 or > 2) return false;
                transform = OfficeTransform.Translate(values[0], values.Count == 2 ? values[1] : 0D);
                return true;
            case "scale":
                if (values.Count is < 1 or > 2) return false;
                transform = OfficeTransform.Scale(values[0], values.Count == 2 ? values[1] : values[0]);
                return true;
            case "rotate":
                if (values.Count != 1 && values.Count != 3) return false;
                transform = values.Count == 1
                    ? OfficeTransform.RotateDegrees(values[0])
                    : OfficeTransform.RotateDegrees(values[0], values[1], values[2]);
                return true;
            case "skewx":
                if (values.Count != 1) return false;
                double tangentX = Math.Tan(values[0] * Math.PI / 180D);
                if (!IsFinite(tangentX)) return false;
                transform = new OfficeTransform(1D, 0D, tangentX, 1D, 0D, 0D);
                return true;
            case "skewy":
                if (values.Count != 1) return false;
                double tangentY = Math.Tan(values[0] * Math.PI / 180D);
                if (!IsFinite(tangentY)) return false;
                transform = new OfficeTransform(1D, tangentY, 0D, 1D, 0D, 0D);
                return true;
            default:
                return false;
        }
    }

    private static bool TryParseNumbers(string value, out IReadOnlyList<double> numbers) {
        var result = new List<double>();
        numbers = result;
        int index = 0;
        while (SkipSeparators(value, ref index)) {
            int start = index;
            if (value[index] is '+' or '-') index++;
            bool digits = false;
            while (index < value.Length && char.IsDigit(value[index])) {
                digits = true;
                index++;
            }
            if (index < value.Length && value[index] == '.') {
                index++;
                while (index < value.Length && char.IsDigit(value[index])) {
                    digits = true;
                    index++;
                }
            }
            if (!digits) return false;
            if (index < value.Length && (value[index] is 'e' or 'E')) {
                int exponent = index++;
                if (index < value.Length && (value[index] is '+' or '-')) index++;
                int exponentDigits = index;
                while (index < value.Length && char.IsDigit(value[index])) index++;
                if (index == exponentDigits) index = exponent;
            }
            if (!double.TryParse(value.Substring(start, index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                || !IsFinite(number)) return false;
            result.Add(number);
        }
        return result.Count > 0;
    }

    private static bool SkipSeparators(string value, ref int index) {
        while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) index++;
        return index < value.Length;
    }

    private static void SkipWhitespace(string value, ref int index) {
        while (index < value.Length && char.IsWhiteSpace(value[index])) index++;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
