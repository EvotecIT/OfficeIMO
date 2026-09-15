using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Html.Css;

internal static class HtmlCssTypedValueParsers {
    internal static bool TryParseNumeric(
        string source,
        IReadOnlyList<HtmlCssToken> tokens,
        CancellationToken cancellationToken,
        out HtmlCssNumericValue? value) {
        var significant = tokens.Where(token => token.Kind != HtmlCssTokenKind.Whitespace
            && token.Kind != HtmlCssTokenKind.Comment && token.Kind != HtmlCssTokenKind.EndOfFile).ToList();
        if (significant.Count != 1 && (significant.Count == 0 || significant[0].Kind != HtmlCssTokenKind.Function
            || !IsMathFunction(significant[0].Value))) {
            value = null;
            return false;
        }
        var parser = new NumericParser(source, significant, cancellationToken);
        if (!parser.TryParse(out Numeric result) || !result.IsFinite) {
            value = null;
            return false;
        }
        value = new HtmlCssNumericValue(result.Type, result.Value, parser.UsedMathFunction);
        return true;
    }

    private static bool IsMathFunction(string? value) {
        string name = (value ?? string.Empty).ToLowerInvariant();
        return name == "calc" || name == "min" || name == "max" || name == "clamp";
    }

    internal static bool TryParseColorFunction(
        string source,
        IReadOnlyList<HtmlCssToken> tokens,
        CancellationToken cancellationToken,
        out HtmlCssColorFunctionValue? value) {
        value = null;
        var significant = tokens.Where(token => token.Kind != HtmlCssTokenKind.Whitespace
            && token.Kind != HtmlCssTokenKind.Comment && token.Kind != HtmlCssTokenKind.EndOfFile).ToList();
        if (significant.Count < 3 || significant[0].Kind != HtmlCssTokenKind.Function
            || significant[significant.Count - 1].Kind != HtmlCssTokenKind.CloseParenthesis) return false;
        string name = (significant[0].Value ?? string.Empty).ToLowerInvariant();
        HtmlCssColorFunctionKind kind;
        if (name == "rgb" || name == "rgba") kind = HtmlCssColorFunctionKind.Rgb;
        else if (name == "hsl" || name == "hsla") kind = HtmlCssColorFunctionKind.Hsl;
        else if (name == "hwb") kind = HtmlCssColorFunctionKind.Hwb;
        else return false;
        cancellationToken.ThrowIfCancellationRequested();
        var body = significant.Skip(1).Take(significant.Count - 2).ToList();
        if (body.Any(token => token.Kind == HtmlCssTokenKind.Function
            || token.Kind == HtmlCssTokenKind.OpenParenthesis || token.Kind == HtmlCssTokenKind.CloseParenthesis)) return false;
        bool commaSyntax = body.Any(token => token.Kind == HtmlCssTokenKind.Comma);
        if (kind == HtmlCssColorFunctionKind.Hwb && commaSyntax) return false;
        List<List<HtmlCssToken>> channels;
        List<HtmlCssToken>? alphaTokens;
        if (commaSyntax) {
            if (body.Any(token => token.Kind == HtmlCssTokenKind.Delimiter && token.Value == "/")) return false;
            List<List<HtmlCssToken>> parts = Split(body, HtmlCssTokenKind.Comma, null);
            if (parts.Count != 3 && parts.Count != 4 || parts.Any(part => part.Count != 1)) return false;
            channels = parts.Take(3).ToList();
            alphaTokens = parts.Count == 4 ? parts[3] : null;
        } else {
            int slash = body.FindIndex(token => token.Kind == HtmlCssTokenKind.Delimiter && token.Value == "/");
            if (slash >= 0 && body.Skip(slash + 1).Any(token => token.Kind == HtmlCssTokenKind.Delimiter && token.Value == "/")) return false;
            List<HtmlCssToken> channelTokens = slash < 0 ? body : body.Take(slash).ToList();
            List<HtmlCssToken> alpha = slash < 0 ? new List<HtmlCssToken>() : body.Skip(slash + 1).ToList();
            if (channelTokens.Count != 3 || alpha.Count > 1) return false;
            channels = channelTokens.Select(token => new List<HtmlCssToken> { token }).ToList();
            alphaTokens = alpha.Count == 0 ? null : alpha;
        }
        var components = new List<HtmlCssColorComponent>(3);
        for (int index = 0; index < 3; index++) {
            HtmlCssColorComponent? component;
            if (!TryColorComponent(source, channels[index][0], kind, index, out component)) return false;
            components.Add(component!);
        }
        if (commaSyntax && kind == HtmlCssColorFunctionKind.Rgb
            && components.Any(component => component.Kind == HtmlCssColorComponentKind.Percentage)
            && components.Any(component => component.Kind == HtmlCssColorComponentKind.Number)) return false;
        if (commaSyntax && kind == HtmlCssColorFunctionKind.Hsl
            && components.Skip(1).Any(component => component.Kind != HtmlCssColorComponentKind.Percentage)) return false;
        HtmlCssColorComponent alphaComponent = new HtmlCssColorComponent(HtmlCssColorComponentKind.Number, 1D);
        if (alphaTokens != null && !TryAlpha(source, alphaTokens[0], out alphaComponent)) return false;
        if (commaSyntax && (components.Any(component => component.Kind == HtmlCssColorComponentKind.None)
            || alphaComponent.Kind == HtmlCssColorComponentKind.None)) return false;
        value = new HtmlCssColorFunctionValue(kind,
            new ReadOnlyCollection<HtmlCssColorComponent>(components), alphaComponent, commaSyntax);
        return true;
    }

    private static bool TryColorComponent(string source, HtmlCssToken token, HtmlCssColorFunctionKind kind, int index, out HtmlCssColorComponent? component) {
        component = null;
        if (token.Kind == HtmlCssTokenKind.Identifier && string.Equals(token.Value, "none", StringComparison.OrdinalIgnoreCase)) {
            component = new HtmlCssColorComponent(HtmlCssColorComponentKind.None, null); return true;
        }
        if (kind == HtmlCssColorFunctionKind.Rgb) return TryNumberOrPercentage(source, token, out component);
        if (index == 0) return TryAngle(source, token, out component);
        return TryNumberOrPercentage(source, token, out component);
    }

    private static bool TryAlpha(string source, HtmlCssToken token, out HtmlCssColorComponent component) {
        if (token.Kind == HtmlCssTokenKind.Identifier && string.Equals(token.Value, "none", StringComparison.OrdinalIgnoreCase)) {
            component = new HtmlCssColorComponent(HtmlCssColorComponentKind.None, null); return true;
        }
        if (TryNumberOrPercentage(source, token, out HtmlCssColorComponent? parsed)) {
            component = parsed!; return true;
        }
        component = new HtmlCssColorComponent(HtmlCssColorComponentKind.Number, 1D);
        return false;
    }

    private static bool TryNumberOrPercentage(string source, HtmlCssToken token, out HtmlCssColorComponent? component) {
        component = null;
        bool percentage = token.Kind == HtmlCssTokenKind.Percentage;
        if (!percentage && token.Kind != HtmlCssTokenKind.Number) return false;
        if (!TryFiniteNumber(TokenNumberText(source, token), out double number)) return false;
        component = new HtmlCssColorComponent(percentage ? HtmlCssColorComponentKind.Percentage : HtmlCssColorComponentKind.Number, number);
        return true;
    }

    private static bool TryAngle(string source, HtmlCssToken token, out HtmlCssColorComponent? component) {
        component = null;
        if (token.Kind == HtmlCssTokenKind.Identifier && string.Equals(token.Value, "none", StringComparison.OrdinalIgnoreCase)) {
            component = new HtmlCssColorComponent(HtmlCssColorComponentKind.None, null); return true;
        }
        double multiplier = 1D;
        if (token.Kind == HtmlCssTokenKind.Dimension) {
            string unit = (token.Value ?? string.Empty).ToLowerInvariant();
            if (unit == "grad") multiplier = 0.9D;
            else if (unit == "turn") multiplier = 360D;
            else if (unit == "rad") multiplier = 180D / Math.PI;
            else if (unit != "deg") return false;
        } else if (token.Kind != HtmlCssTokenKind.Number) return false;
        string text = token.GetText(source);
        if (token.Kind == HtmlCssTokenKind.Dimension) text = text.Substring(0, text.Length - (token.Value ?? string.Empty).Length);
        if (!TryFiniteNumber(text, out double number)) return false;
        double degrees = number * multiplier % 360D;
        if (degrees < 0D) degrees += 360D;
        component = new HtmlCssColorComponent(HtmlCssColorComponentKind.Angle, degrees);
        return true;
    }

    private static string TokenNumberText(string source, HtmlCssToken token) {
        string text = token.GetText(source);
        return token.Kind == HtmlCssTokenKind.Percentage ? text.Substring(0, text.Length - 1) : text;
    }

    private static bool TryFiniteNumber(string text, out double value) =>
        double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
        && !double.IsNaN(value) && !double.IsInfinity(value);

    private static List<List<HtmlCssToken>> Split(IReadOnlyList<HtmlCssToken> tokens, HtmlCssTokenKind kind, string? delimiter) {
        var result = new List<List<HtmlCssToken>> { new List<HtmlCssToken>() };
        foreach (HtmlCssToken token in tokens) {
            if (token.Kind == kind && (delimiter == null || token.Value == delimiter)) result.Add(new List<HtmlCssToken>());
            else result[result.Count - 1].Add(token);
        }
        return result;
    }

    private sealed class NumericParser {
        private readonly string _source;
        private readonly IReadOnlyList<HtmlCssToken> _tokens;
        private readonly CancellationToken _cancellation;
        private int _position;
        private int _depth;
        internal NumericParser(string source, IReadOnlyList<HtmlCssToken> tokens, CancellationToken cancellation) {
            _source = source; _tokens = tokens; _cancellation = cancellation;
        }
        internal bool UsedMathFunction { get; private set; }
        internal bool TryParse(out Numeric result) {
            if (!TryExpression(out result) || _position != _tokens.Count) return false;
            return result.IsFinite;
        }
        private bool TryExpression(out Numeric result) {
            if (!TryProduct(out result)) return false;
            while (IsDelimiter("+") || IsDelimiter("-")) {
                if (!HasRequiredBinaryWhitespace()) return false;
                bool subtract = Current.Value == "-"; _position++; RecordOperation();
                if (!TryProduct(out Numeric right) || result.Type != right.Type) return false;
                result = new Numeric(subtract ? result.Value - right.Value : result.Value + right.Value, result.Type);
            }
            return result.IsFinite;
        }
        private bool TryProduct(out Numeric result) {
            if (!TryPrimary(out result)) return false;
            while (IsDelimiter("*") || IsDelimiter("/")) {
                bool divide = Current.Value == "/"; _position++; RecordOperation();
                if (!TryPrimary(out Numeric right)) return false;
                if (divide) {
                    if (right.Value == 0D) return false;
                    if (right.Type == HtmlCssNumericType.Percentage) {
                        if (result.Type != HtmlCssNumericType.Percentage) return false;
                        result = new Numeric(result.Value / right.Value, HtmlCssNumericType.Number);
                    } else result = new Numeric(result.Value / right.Value, result.Type);
                } else {
                    if (result.Type == HtmlCssNumericType.Percentage && right.Type == HtmlCssNumericType.Percentage) return false;
                    HtmlCssNumericType type = result.Type == HtmlCssNumericType.Percentage || right.Type == HtmlCssNumericType.Percentage
                        ? HtmlCssNumericType.Percentage : HtmlCssNumericType.Number;
                    result = new Numeric(result.Value * right.Value, type);
                }
                if (!result.IsFinite) return false;
            }
            return true;
        }
        private bool TryPrimary(out Numeric result) {
            _cancellation.ThrowIfCancellationRequested();
            if (Current.Kind == HtmlCssTokenKind.Number || Current.Kind == HtmlCssTokenKind.Percentage) {
                bool percentage = Current.Kind == HtmlCssTokenKind.Percentage;
                string text = TokenNumberText(_source, Current);
                _position++;
                if (TryFiniteNumber(text, out double number)) {
                    result = new Numeric(number, percentage ? HtmlCssNumericType.Percentage : HtmlCssNumericType.Number); return true;
                }
            } else if (Current.Kind == HtmlCssTokenKind.OpenParenthesis) {
                if (!Enter()) { result = default; return false; }
                _position++;
                if (TryExpression(out result) && Current.Kind == HtmlCssTokenKind.CloseParenthesis) { _position++; _depth--; return true; }
                _depth--;
            } else if (Current.Kind == HtmlCssTokenKind.Function) {
                return TryFunction(out result);
            }
            result = default;
            return false;
        }
        private bool TryFunction(out Numeric result) {
            string name = (Current.Value ?? string.Empty).ToLowerInvariant();
            if (name != "calc" && name != "min" && name != "max" && name != "clamp") { result = default; return false; }
            if (!Enter()) { result = default; return false; }
            UsedMathFunction = true; _position++; RecordOperation();
            var args = new List<Numeric>();
            if (!TryExpression(out Numeric first)) { _depth--; result = default; return false; }
            args.Add(first);
            while (Current.Kind == HtmlCssTokenKind.Comma) {
                _position++; RecordOperation();
                if (!TryExpression(out Numeric next)) { _depth--; result = default; return false; }
                args.Add(next);
            }
            if (Current.Kind != HtmlCssTokenKind.CloseParenthesis) { _depth--; result = default; return false; }
            _position++;
            _depth--;
            if (name == "calc") { result = first; return args.Count == 1; }
            int required = name == "clamp" ? 3 : 1;
            if (args.Count < required || name == "clamp" && args.Count != 3 || args.Any(arg => arg.Type != first.Type)) {
                result = default; return false;
            }
            double value = first.Value;
            if (name == "min") foreach (Numeric arg in args) value = Math.Min(value, arg.Value);
            else if (name == "max") foreach (Numeric arg in args) value = Math.Max(value, arg.Value);
            else value = Math.Max(args[0].Value, Math.Min(args[1].Value, args[2].Value));
            result = new Numeric(value, first.Type);
            return result.IsFinite;
        }
        private void RecordOperation() {
            _cancellation.ThrowIfCancellationRequested();
        }
        private bool HasRequiredBinaryWhitespace() {
            if (_position <= 0 || _position + 1 >= _tokens.Count) return false;
            HtmlCssToken previous = _tokens[_position - 1];
            HtmlCssToken next = _tokens[_position + 1];
            return GapContainsWhitespace(previous.Offset + previous.Length, Current.Offset)
                && GapContainsWhitespace(Current.Offset + Current.Length, next.Offset);
        }
        private bool GapContainsWhitespace(int start, int end) {
            for (int index = start; index < end; index++) {
                if (_source[index] == '/' && index + 1 < end && _source[index + 1] == '*') {
                    int close = _source.IndexOf("*/", index + 2, StringComparison.Ordinal);
                    if (close < 0 || close >= end) return false;
                    index = close + 1;
                    continue;
                }
                char value = _source[index];
                if (value == ' ' || value == '\t' || value == '\r' || value == '\n' || value == '\f') return true;
            }
            return false;
        }
        private bool Enter() => ++_depth <= 128;
        private HtmlCssToken Current => _position < _tokens.Count ? _tokens[_position] : default;
        private bool IsDelimiter(string value) => _position < _tokens.Count && Current.Kind == HtmlCssTokenKind.Delimiter && Current.Value == value;
    }

    private readonly struct Numeric {
        internal Numeric(double value, HtmlCssNumericType type) { Value = value; Type = type; }
        internal double Value { get; }
        internal HtmlCssNumericType Type { get; }
        internal bool IsFinite => !double.IsNaN(Value) && !double.IsInfinity(Value);
    }
}
