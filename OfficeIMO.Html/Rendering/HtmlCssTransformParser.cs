using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssTransformParser {
    internal static bool TryParse(
        string transformValue,
        string transformOriginValue,
        double boxX,
        double boxY,
        double boxWidth,
        double boxHeight,
        double fontSize,
        double rootFontSize,
        out OfficeTransform transform,
        out string detail) {
        transform = OfficeTransform.Identity;
        detail = string.Empty;
        string value = transformValue.Trim().ToLowerInvariant();
        if (value.Length == 0 || value == "none") return true;
        if (!TryParseFunctionList(value, boxWidth, boxHeight, fontSize, rootFontSize, out OfficeTransform functions, out detail)) return false;
        if (!TryResolveOrigin(transformOriginValue, boxWidth, boxHeight, fontSize, rootFontSize, out double originX, out double originY)) {
            detail = "transform-origin=" + transformOriginValue.Trim();
            return false;
        }
        originX += boxX;
        originY += boxY;
        if (!IsFinite(originX) || !IsFinite(originY)) {
            detail = "transform-origin=" + transformOriginValue.Trim();
            return false;
        }
        if (!TryThen(OfficeTransform.Translate(-originX, -originY), functions, out OfficeTransform centered)
            || !TryThen(centered, OfficeTransform.Translate(originX, originY), out transform)) {
            detail = "transform=" + value;
            return false;
        }
        return true;
    }

    internal static bool IsSupportedTransformSyntax(string value) =>
        TryParseFunctionList(value.Trim().ToLowerInvariant(), 100D, 100D, 16D, 16D, out _, out _);

    internal static bool IsSupportedOriginSyntax(string value) =>
        TryResolveOrigin(value, 100D, 100D, 16D, 16D, out _, out _);

    private static bool TryParseFunctionList(
        string value,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out OfficeTransform transform,
        out string detail) {
        transform = OfficeTransform.Identity;
        detail = string.Empty;
        if (value == "none") return true;
        int index = 0;
        bool found = false;
        while (index < value.Length) {
            while (index < value.Length && char.IsWhiteSpace(value[index])) index++;
            if (index >= value.Length) break;
            int nameStart = index;
            while (index < value.Length && (char.IsLetterOrDigit(value[index]) || value[index] == '-')) index++;
            if (index == nameStart || index >= value.Length || value[index] != '(') {
                detail = "transform=" + value;
                return false;
            }
            string name = value.Substring(nameStart, index - nameStart);
            int argumentsStart = ++index;
            int depth = 1;
            while (index < value.Length && depth > 0) {
                if (value[index] == '(') depth++;
                else if (value[index] == ')') depth--;
                index++;
            }
            if (depth != 0) {
                detail = "transform=" + value;
                return false;
            }
            string arguments = value.Substring(argumentsStart, index - argumentsStart - 1);
            if (!TryParseFunction(name, arguments, width, height, fontSize, rootFontSize, out OfficeTransform function)) {
                detail = name + "(" + arguments + ")";
                return false;
            }
            if (!TryThen(function, transform, out transform)) {
                detail = name + "(" + arguments + ")";
                return false;
            }
            found = true;
        }
        if (!found) detail = "transform=" + value;
        return found;
    }

    private static bool TryParseFunction(
        string name,
        string arguments,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out OfficeTransform transform) {
        transform = OfficeTransform.Identity;
        IReadOnlyList<string> values = SplitArguments(arguments);
        switch (name) {
            case "matrix":
                if (values.Count != 6 || !TryNumbers(values, out double[]? matrix)) return false;
                transform = new OfficeTransform(matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]);
                return true;
            case "translate":
                double translateY = 0D;
                if (values.Count < 1 || values.Count > 2
                    || !TryLength(values[0], width, fontSize, rootFontSize, out double translateX)
                    || values.Count == 2 && !TryLength(values[1], height, fontSize, rootFontSize, out translateY)) return false;
                transform = OfficeTransform.Translate(translateX, values.Count == 2 ? translateY : 0D);
                return true;
            case "translatex":
                if (values.Count != 1 || !TryLength(values[0], width, fontSize, rootFontSize, out double x)) return false;
                transform = OfficeTransform.Translate(x, 0D);
                return true;
            case "translatey":
                if (values.Count != 1 || !TryLength(values[0], height, fontSize, rootFontSize, out double y)) return false;
                transform = OfficeTransform.Translate(0D, y);
                return true;
            case "scale":
                double scaleY = 1D;
                if (values.Count < 1 || values.Count > 2 || !TryNumber(values[0], out double scaleX)
                    || values.Count == 2 && !TryNumber(values[1], out scaleY)) return false;
                transform = OfficeTransform.Scale(scaleX, values.Count == 2 ? scaleY : scaleX);
                return true;
            case "scalex":
                if (values.Count != 1 || !TryNumber(values[0], out double sx)) return false;
                transform = OfficeTransform.Scale(sx, 1D);
                return true;
            case "scaley":
                if (values.Count != 1 || !TryNumber(values[0], out double sy)) return false;
                transform = OfficeTransform.Scale(1D, sy);
                return true;
            case "rotate":
                if (values.Count != 1 || !TryAngle(values[0], out double rotation)) return false;
                transform = OfficeTransform.RotateDegrees(rotation);
                return true;
            case "skew":
                double skewY = 0D;
                if (values.Count < 1 || values.Count > 2 || !TryAngle(values[0], out double skewX)
                    || values.Count == 2 && !TryAngle(values[1], out skewY)) return false;
                return TryCreateSkew(skewX, values.Count == 2 ? skewY : 0D, out transform);
            case "skewx":
                if (values.Count != 1 || !TryAngle(values[0], out double xAngle)) return false;
                return TryCreateSkew(xAngle, 0D, out transform);
            case "skewy":
                if (values.Count != 1 || !TryAngle(values[0], out double yAngle)) return false;
                return TryCreateSkew(0D, yAngle, out transform);
            default:
                return false;
        }
    }

    private static bool TryCreateSkew(double xDegrees, double yDegrees, out OfficeTransform transform) {
        double xRadians = Math.IEEERemainder(xDegrees, 180D) * Math.PI / 180D;
        double yRadians = Math.IEEERemainder(yDegrees, 180D) * Math.PI / 180D;
        double x = Math.Tan(xRadians);
        double y = Math.Tan(yRadians);
        if (!IsFinite(x) || !IsFinite(y)) {
            transform = OfficeTransform.Identity;
            return false;
        }
        transform = new OfficeTransform(1D, y, x, 1D, 0D, 0D);
        return true;
    }

    private static bool TryThen(OfficeTransform first, OfficeTransform next, out OfficeTransform transform) {
        double m11 = next.M11 * first.M11 + next.M21 * first.M12;
        double m12 = next.M12 * first.M11 + next.M22 * first.M12;
        double m21 = next.M11 * first.M21 + next.M21 * first.M22;
        double m22 = next.M12 * first.M21 + next.M22 * first.M22;
        double offsetX = next.M11 * first.OffsetX + next.M21 * first.OffsetY + next.OffsetX;
        double offsetY = next.M12 * first.OffsetX + next.M22 * first.OffsetY + next.OffsetY;
        if (!IsFinite(m11) || !IsFinite(m12) || !IsFinite(m21) || !IsFinite(m22)
            || !IsFinite(offsetX) || !IsFinite(offsetY)) {
            transform = OfficeTransform.Identity;
            return false;
        }
        transform = new OfficeTransform(m11, m12, m21, m22, offsetX, offsetY);
        return true;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static IReadOnlyList<string> SplitArguments(string value) {
        IReadOnlyList<string> commas = HtmlRenderCssValues.SplitTopLevelCommas(value);
        if (commas.Count > 1) return commas.Select(item => item.Trim()).Where(item => item.Length > 0).ToList();
        return HtmlRenderCssValues.SplitWhitespace(value);
    }

    private static bool TryNumbers(IReadOnlyList<string> values, out double[] numbers) {
        numbers = new double[values.Count];
        for (int index = 0; index < values.Count; index++) {
            if (!TryNumber(values[index], out numbers[index])) return false;
        }
        return true;
    }

    private static bool TryNumber(string value, out double number) =>
        double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number)
        && !double.IsNaN(number)
        && !double.IsInfinity(number);

    private static bool TryLength(string value, double reference, double fontSize, double rootFontSize, out double length) =>
        HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, out length)
        && !double.IsNaN(length)
        && !double.IsInfinity(length);

    private static bool TryAngle(string value, out double degrees) {
        degrees = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        double factor;
        string number;
        if (normalized.EndsWith("deg", StringComparison.Ordinal)) {
            factor = 1D;
            number = normalized.Substring(0, normalized.Length - 3);
        } else if (normalized.EndsWith("grad", StringComparison.Ordinal)) {
            factor = 0.9D;
            number = normalized.Substring(0, normalized.Length - 4);
        } else if (normalized.EndsWith("rad", StringComparison.Ordinal)) {
            factor = 180D / Math.PI;
            number = normalized.Substring(0, normalized.Length - 3);
        } else if (normalized.EndsWith("turn", StringComparison.Ordinal)) {
            factor = 360D;
            number = normalized.Substring(0, normalized.Length - 4);
        } else if (normalized == "0") {
            return true;
        } else {
            return false;
        }
        if (!TryNumber(number, out double parsed)) return false;
        degrees = parsed * factor;
        return !double.IsNaN(degrees) && !double.IsInfinity(degrees);
    }

    private static bool TryResolveOrigin(
        string value,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out double x,
        out double y) {
        x = width / 2D;
        y = height / 2D;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(string.IsNullOrWhiteSpace(value) ? "50% 50%" : value.Trim().ToLowerInvariant());
        if (parts.Count == 0 || parts.Count > 2) return false;
        string first = parts[0];
        string second = parts.Count == 2 ? parts[1] : "center";
        if (IsVerticalKeyword(first) && IsHorizontalKeyword(second)) {
            string swap = first;
            first = second;
            second = swap;
        } else if (parts.Count == 1 && IsVerticalKeyword(first)) {
            second = first;
            first = "center";
        }
        return TryOriginAxis(first, width, fontSize, rootFontSize, horizontal: true, out x)
            && TryOriginAxis(second, height, fontSize, rootFontSize, horizontal: false, out y);
    }

    private static bool TryOriginAxis(string value, double reference, double fontSize, double rootFontSize, bool horizontal, out double result) {
        result = reference / 2D;
        if (value == "center") return true;
        if (horizontal && value == "left" || !horizontal && value == "top") {
            result = 0D;
            return true;
        }
        if (horizontal && value == "right" || !horizontal && value == "bottom") {
            result = reference;
            return true;
        }
        if (horizontal && IsVerticalKeyword(value) || !horizontal && IsHorizontalKeyword(value)) return false;
        return TryLength(value, reference, fontSize, rootFontSize, out result);
    }

    private static bool IsHorizontalKeyword(string value) => value == "left" || value == "center" || value == "right";
    private static bool IsVerticalKeyword(string value) => value == "top" || value == "center" || value == "bottom";
}
