using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Parses the bounded OfficeIMO-owned CSS length, percentage, and arithmetic subset.</summary>
public static class HtmlCssMathParser {
    /// <summary>Parses a standalone value accepted by a &lt;length-percentage&gt; property position.</summary>
    public static HtmlCssMathParseResult ParseLengthPercentage(
        string source,
        HtmlCssMathOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        HtmlCssMathOptions effective = (options ?? new HtmlCssMathOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Length > effective.MaxInputCharacters)
            throw new HtmlCssMathLimitException(nameof(effective.MaxInputCharacters), source.Length, effective.MaxInputCharacters);
        string authored = source.Trim();
        if (TrySimpleLiteral(authored, out HtmlCssMathParseStatus simpleStatus,
                out HtmlCssMathExpression? simpleExpression))
            return Finish(authored, simpleStatus, simpleExpression);
        IReadOnlyList<HtmlCssToken> tokenized;
        try {
            tokenized = HtmlCssTokenizer.Tokenize(authored,
                new HtmlCssTokenizationOptions { MaxInputCharacters = effective.MaxInputCharacters, MaxTokens = effective.MaxTokens },
                cancellationToken);
        } catch (HtmlCssTokenizationLimitException exception) when (exception.LimitName == nameof(HtmlCssTokenizationOptions.MaxTokens)) {
            throw new HtmlCssMathLimitException(nameof(effective.MaxTokens), exception.Actual, exception.Maximum);
        }
        var tokens = tokenized.Where(token => token.Kind != HtmlCssTokenKind.Whitespace
            && token.Kind != HtmlCssTokenKind.Comment && token.Kind != HtmlCssTokenKind.EndOfFile).ToList();
        if (tokens.Count == 0) return Result(authored, HtmlCssMathParseStatus.InvalidSyntax, null);

        var parser = new Parser(authored, tokens, effective, cancellationToken);
        if (!parser.TryParse(out HtmlCssMathExpression? expression))
            return Result(authored, parser.FailureStatus, null);
        return Finish(authored, HtmlCssMathParseStatus.Parsed, expression);
    }

    internal static HtmlCssMathParseResult ParseLengthPercentage(
        string source,
        IReadOnlyList<HtmlCssToken> significantTokens,
        CancellationToken cancellationToken) {
        HtmlCssMathOptions effective = new HtmlCssMathOptions();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Length > effective.MaxInputCharacters)
            throw new HtmlCssMathLimitException(nameof(effective.MaxInputCharacters), source.Length, effective.MaxInputCharacters);
        if (significantTokens.Count > effective.MaxTokens)
            throw new HtmlCssMathLimitException(nameof(effective.MaxTokens), significantTokens.Count, effective.MaxTokens);
        if (TrySimpleLiteral(source, out HtmlCssMathParseStatus simpleStatus,
                out HtmlCssMathExpression? simpleExpression))
            return Finish(source, simpleStatus, simpleExpression);
        if (significantTokens.Count == 0) return Result(source, HtmlCssMathParseStatus.InvalidSyntax, null);
        var parser = new Parser(source, significantTokens, effective, cancellationToken);
        return parser.TryParse(out HtmlCssMathExpression? expression)
            ? Finish(source, HtmlCssMathParseStatus.Parsed, expression)
            : Result(source, parser.FailureStatus, null);
    }

    private static HtmlCssMathParseResult Finish(
        string authored,
        HtmlCssMathParseStatus status,
        HtmlCssMathExpression? expression) {
        if (status != HtmlCssMathParseStatus.Parsed || expression == null)
            return Result(authored, status, null);
        if (expression!.Kind == HtmlCssMathExpressionKind.Literal
            && expression.Type == HtmlCssNumericType.Number && expression.Value == 0D)
            expression = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal, HtmlCssNumericType.Length, "0", 0D);
        if (expression!.Type != HtmlCssNumericType.Length
            && expression.Type != HtmlCssNumericType.Percentage
            && expression.Type != HtmlCssNumericType.LengthPercentage)
            return Result(authored, HtmlCssMathParseStatus.IncompatibleTypes, null);
        return Result(authored, HtmlCssMathParseStatus.Parsed, expression);
    }

    private static bool TrySimpleLiteral(
        string text,
        out HtmlCssMathParseStatus status,
        out HtmlCssMathExpression? expression) {
        status = HtmlCssMathParseStatus.InvalidSyntax;
        expression = null;
        if (text.Length == 0 || !TryNumericPrefix(text, out double number, out int numberLength)) return false;
        string suffix = text.Substring(numberLength);
        string canonicalNumber = number.ToString("R", CultureInfo.InvariantCulture);
        if (suffix.Length == 0) {
            status = HtmlCssMathParseStatus.Parsed;
            expression = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
                HtmlCssNumericType.Number, canonicalNumber, number);
            return true;
        }
        if (suffix == "%") {
            status = HtmlCssMathParseStatus.Parsed;
            expression = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
                HtmlCssNumericType.Percentage, canonicalNumber + "%", number);
            return true;
        }
        if (suffix.Any(character => character < 'A' || character > 'Z' && character < 'a' || character > 'z'))
            return false;
        if (!TryUnit(suffix, out HtmlCssLengthUnit unit, out string? canonicalUnit)) {
            status = HtmlCssMathParseStatus.UnsupportedValue;
            return true;
        }
        status = HtmlCssMathParseStatus.Parsed;
        expression = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
            HtmlCssNumericType.Length, canonicalNumber + canonicalUnit, number, unit);
        return true;
    }

    private static HtmlCssMathParseResult Result(
        string authored,
        HtmlCssMathParseStatus status,
        HtmlCssMathExpression? expression) => new HtmlCssMathParseResult(authored, status, expression);

    private sealed class Parser {
        private readonly string _source;
        private readonly IReadOnlyList<HtmlCssToken> _tokens;
        private readonly HtmlCssMathOptions _options;
        private readonly CancellationToken _cancellation;
        private int _position;
        private int _depth;
        private int _operations;

        internal Parser(
            string source,
            IReadOnlyList<HtmlCssToken> tokens,
            HtmlCssMathOptions options,
            CancellationToken cancellation) {
            _source = source;
            _tokens = tokens;
            _options = options;
            _cancellation = cancellation;
        }

        internal HtmlCssMathParseStatus FailureStatus { get; private set; } = HtmlCssMathParseStatus.InvalidSyntax;

        internal bool TryParse(out HtmlCssMathExpression? result) {
            _cancellation.ThrowIfCancellationRequested();
            bool rootIsLiteral = _tokens.Count == 1 && IsLiteral(Current.Kind);
            bool rootIsFunction = Current.Kind == HtmlCssTokenKind.Function && IsMathFunction(Current.Value);
            if (!rootIsLiteral && !rootIsFunction) {
                if (Current.Kind == HtmlCssTokenKind.Function) FailureStatus = HtmlCssMathParseStatus.UnsupportedValue;
                result = null;
                return false;
            }
            if (!TryExpression(out result) || _position != _tokens.Count) {
                result = null;
                return false;
            }
            return true;
        }

        private bool TryExpression(out HtmlCssMathExpression? result) {
            if (!TryProduct(out result)) return false;
            while (IsDelimiter("+") || IsDelimiter("-")) {
                if (!HasRequiredBinaryWhitespace()) { result = null; return false; }
                bool subtract = Current.Value == "-";
                _position++;
                RecordOperation();
                if (!TryProduct(out HtmlCssMathExpression? right)) { result = null; return false; }
                if (!TryAddType(result!.Type, right!.Type, out HtmlCssNumericType type)) {
                    FailureStatus = HtmlCssMathParseStatus.IncompatibleTypes;
                    result = null;
                    return false;
                }
                HtmlCssMathExpressionKind kind = subtract
                    ? HtmlCssMathExpressionKind.Subtract : HtmlCssMathExpressionKind.Add;
                result = Node(kind, type, BinaryCanonical(kind, result, right), result, right);
            }
            return true;
        }

        private bool TryProduct(out HtmlCssMathExpression? result) {
            if (!TryPrimary(out result)) return false;
            while (IsDelimiter("*") || IsDelimiter("/")) {
                bool divide = Current.Value == "/";
                _position++;
                RecordOperation();
                if (!TryPrimary(out HtmlCssMathExpression? right)) { result = null; return false; }
                if (!TryProductType(result!.Type, right!.Type, divide, out HtmlCssNumericType type)) {
                    FailureStatus = HtmlCssMathParseStatus.IncompatibleTypes;
                    result = null;
                    return false;
                }
                HtmlCssMathExpressionKind kind = divide
                    ? HtmlCssMathExpressionKind.Divide : HtmlCssMathExpressionKind.Multiply;
                result = Node(kind, type, BinaryCanonical(kind, result, right), result, right);
            }
            return true;
        }

        private bool TryPrimary(out HtmlCssMathExpression? result) {
            _cancellation.ThrowIfCancellationRequested();
            if (AtEnd) { result = null; return false; }
            if (IsLiteral(Current.Kind)) return TryLiteral(out result);
            if (Current.Kind == HtmlCssTokenKind.OpenParenthesis) {
                EnterNesting();
                _position++;
                bool parsed = TryExpression(out result) && !AtEnd && Current.Kind == HtmlCssTokenKind.CloseParenthesis;
                if (parsed) _position++;
                ExitNesting();
                if (!parsed) result = null;
                return parsed;
            }
            if (Current.Kind == HtmlCssTokenKind.Function) return TryFunction(out result);
            result = null;
            return false;
        }

        private bool TryLiteral(out HtmlCssMathExpression? result) {
            HtmlCssToken token = Current;
            _position++;
            if (!TryTokenNumber(_source, token, out double number)) { result = null; return false; }
            string canonicalNumber = number.ToString("R", CultureInfo.InvariantCulture);
            if (token.Kind == HtmlCssTokenKind.Number) {
                result = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
                    HtmlCssNumericType.Number, canonicalNumber, number);
                return true;
            }
            if (token.Kind == HtmlCssTokenKind.Percentage) {
                result = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
                    HtmlCssNumericType.Percentage, canonicalNumber + "%", number);
                return true;
            }
            if (!TryUnit(token.Value, out HtmlCssLengthUnit unit, out string? canonicalUnit)) {
                FailureStatus = HtmlCssMathParseStatus.UnsupportedValue;
                result = null;
                return false;
            }
            result = new HtmlCssMathExpression(HtmlCssMathExpressionKind.Literal,
                HtmlCssNumericType.Length, canonicalNumber + canonicalUnit, number, unit);
            return true;
        }

        private bool TryFunction(out HtmlCssMathExpression? result) {
            string name = (Current.Value ?? string.Empty).ToLowerInvariant();
            if (!IsMathFunction(name)) {
                FailureStatus = HtmlCssMathParseStatus.UnsupportedValue;
                result = null;
                return false;
            }
            EnterNesting();
            _position++;
            RecordOperation();
            var arguments = new List<HtmlCssMathExpression>();
            if (!TryExpression(out HtmlCssMathExpression? first)) {
                ExitNesting();
                result = null;
                return false;
            }
            arguments.Add(first!);
            while (!AtEnd && Current.Kind == HtmlCssTokenKind.Comma) {
                _position++;
                RecordOperation();
                if (arguments.Count >= _options.MaxArguments)
                    throw new HtmlCssMathLimitException(nameof(_options.MaxArguments), arguments.Count + 1, _options.MaxArguments);
                if (!TryExpression(out HtmlCssMathExpression? next)) {
                    ExitNesting();
                    result = null;
                    return false;
                }
                arguments.Add(next!);
            }
            if (AtEnd || Current.Kind != HtmlCssTokenKind.CloseParenthesis) {
                ExitNesting();
                result = null;
                return false;
            }
            _position++;
            ExitNesting();

            if (name == "calc") {
                if (arguments.Count != 1) { result = null; return false; }
                result = Function(HtmlCssMathExpressionKind.Calc, first!.Type, "calc", arguments);
                return true;
            }
            if (name == "clamp" && arguments.Count != 3) { result = null; return false; }
            if ((name == "min" || name == "max") && arguments.Count == 0) { result = null; return false; }
            HtmlCssNumericType type = arguments[0].Type;
            for (int index = 1; index < arguments.Count; index++) {
                if (!TryAddType(type, arguments[index].Type, out type)) {
                    FailureStatus = HtmlCssMathParseStatus.IncompatibleTypes;
                    result = null;
                    return false;
                }
            }
            HtmlCssMathExpressionKind kind = name == "min" ? HtmlCssMathExpressionKind.Min
                : name == "max" ? HtmlCssMathExpressionKind.Max : HtmlCssMathExpressionKind.Clamp;
            result = Function(kind, type, name, arguments);
            return true;
        }

        private HtmlCssMathExpression Function(
            HtmlCssMathExpressionKind kind,
            HtmlCssNumericType type,
            string name,
            IReadOnlyList<HtmlCssMathExpression> arguments) =>
            new HtmlCssMathExpression(kind, type,
                name + "(" + string.Join(", ", arguments.Select(argument => argument.CanonicalText)) + ")",
                children: new ReadOnlyCollection<HtmlCssMathExpression>(arguments.ToList()));

    private static HtmlCssMathExpression Node(
            HtmlCssMathExpressionKind kind,
            HtmlCssNumericType type,
            string canonical,
            params HtmlCssMathExpression[] children) =>
            new HtmlCssMathExpression(kind, type, canonical,
                children: new ReadOnlyCollection<HtmlCssMathExpression>(children));

        private static string BinaryCanonical(
            HtmlCssMathExpressionKind kind,
            HtmlCssMathExpression left,
            HtmlCssMathExpression right) {
            string operation = kind == HtmlCssMathExpressionKind.Add ? " + "
                : kind == HtmlCssMathExpressionKind.Subtract ? " - "
                : kind == HtmlCssMathExpressionKind.Multiply ? " * " : " / ";
            return FormatChild(left, kind, isRight: false) + operation + FormatChild(right, kind, isRight: true);
        }

        private static string FormatChild(
            HtmlCssMathExpression child,
            HtmlCssMathExpressionKind parent,
            bool isRight) {
            int childPrecedence = Precedence(child.Kind);
            int parentPrecedence = Precedence(parent);
            bool samePrecedenceNeedsGrouping = isRight
                && (parent == HtmlCssMathExpressionKind.Subtract
                    && (child.Kind == HtmlCssMathExpressionKind.Add || child.Kind == HtmlCssMathExpressionKind.Subtract)
                    || parent == HtmlCssMathExpressionKind.Divide
                    && (child.Kind == HtmlCssMathExpressionKind.Multiply || child.Kind == HtmlCssMathExpressionKind.Divide));
            return childPrecedence < parentPrecedence || samePrecedenceNeedsGrouping
                ? "(" + child.CanonicalText + ")"
                : child.CanonicalText;
        }

        private static int Precedence(HtmlCssMathExpressionKind kind) {
            if (kind == HtmlCssMathExpressionKind.Add || kind == HtmlCssMathExpressionKind.Subtract) return 1;
            if (kind == HtmlCssMathExpressionKind.Multiply || kind == HtmlCssMathExpressionKind.Divide) return 2;
            return 3;
        }

        private void EnterNesting() {
            _depth++;
            if (_depth > _options.MaxNestingDepth)
                throw new HtmlCssMathLimitException(nameof(_options.MaxNestingDepth), _depth, _options.MaxNestingDepth);
        }

        private void ExitNesting() {
            if (_depth > 0) _depth--;
        }

        private void RecordOperation() {
            _cancellation.ThrowIfCancellationRequested();
            _operations++;
            if (_operations > _options.MaxOperations)
                throw new HtmlCssMathLimitException(nameof(_options.MaxOperations), _operations, _options.MaxOperations);
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

        private bool AtEnd => _position >= _tokens.Count;
        private HtmlCssToken Current => _tokens[_position];
        private bool IsDelimiter(string value) => !AtEnd && Current.Kind == HtmlCssTokenKind.Delimiter && Current.Value == value;
    }

    private static bool TryAddType(HtmlCssNumericType left, HtmlCssNumericType right, out HtmlCssNumericType type) {
        if (left == right) { type = left; return true; }
        if (IsLengthPercentage(left) && IsLengthPercentage(right)) {
            type = HtmlCssNumericType.LengthPercentage;
            return true;
        }
        type = default;
        return false;
    }

    private static bool TryProductType(
        HtmlCssNumericType left,
        HtmlCssNumericType right,
        bool divide,
        out HtmlCssNumericType type) {
        if (divide) {
            if (right == HtmlCssNumericType.Number) { type = left; return true; }
            if (left == right) { type = HtmlCssNumericType.Number; return true; }
            type = default;
            return false;
        }
        if (left == HtmlCssNumericType.Number) { type = right; return true; }
        if (right == HtmlCssNumericType.Number) { type = left; return true; }
        type = default;
        return false;
    }

    private static bool IsLengthPercentage(HtmlCssNumericType type) =>
        type == HtmlCssNumericType.Length || type == HtmlCssNumericType.Percentage || type == HtmlCssNumericType.LengthPercentage;

    private static bool IsLiteral(HtmlCssTokenKind kind) =>
        kind == HtmlCssTokenKind.Number || kind == HtmlCssTokenKind.Percentage || kind == HtmlCssTokenKind.Dimension;

    private static bool IsMathFunction(string? value) {
        string name = (value ?? string.Empty).ToLowerInvariant();
        return name == "calc" || name == "min" || name == "max" || name == "clamp";
    }

    private static bool TryTokenNumber(string source, HtmlCssToken token, out double value) {
        string text = token.GetText(source);
        return TryNumericPrefix(text, out value, out _);
    }

    private static bool TryNumericPrefix(string text, out double value, out int length) {
        length = 0;
        int index = 0;
        if (index < text.Length && (text[index] == '+' || text[index] == '-')) index++;
        int digitStart = index;
        while (index < text.Length && text[index] >= '0' && text[index] <= '9') index++;
        bool hasDigit = index > digitStart;
        if (index < text.Length && text[index] == '.') {
            index++;
            int fractionStart = index;
            while (index < text.Length && text[index] >= '0' && text[index] <= '9') index++;
            if (index == fractionStart) { value = 0D; return false; }
            hasDigit = true;
        }
        if (!hasDigit) { value = 0D; return false; }
        if (index < text.Length && (text[index] == 'e' || text[index] == 'E')) {
            int exponent = index++;
            if (index < text.Length && (text[index] == '+' || text[index] == '-')) index++;
            int exponentDigits = index;
            while (index < text.Length && text[index] >= '0' && text[index] <= '9') index++;
            if (index == exponentDigits) index = exponent;
        }
        length = index;
        return double.TryParse(text.Substring(0, index), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            && !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static bool TryUnit(string? value, out HtmlCssLengthUnit unit, out string? canonical) {
        canonical = (value ?? string.Empty).ToLowerInvariant();
        switch (canonical) {
            case "px": unit = HtmlCssLengthUnit.Px; return true;
            case "pt": unit = HtmlCssLengthUnit.Pt; return true;
            case "pc": unit = HtmlCssLengthUnit.Pc; return true;
            case "in": unit = HtmlCssLengthUnit.In; return true;
            case "cm": unit = HtmlCssLengthUnit.Cm; return true;
            case "mm": unit = HtmlCssLengthUnit.Mm; return true;
            case "q": unit = HtmlCssLengthUnit.Q; return true;
            case "em": unit = HtmlCssLengthUnit.Em; return true;
            case "rem": unit = HtmlCssLengthUnit.Rem; return true;
            case "vw": unit = HtmlCssLengthUnit.Vw; return true;
            case "vh": unit = HtmlCssLengthUnit.Vh; return true;
            case "vmin": unit = HtmlCssLengthUnit.Vmin; return true;
            case "vmax": unit = HtmlCssLengthUnit.Vmax; return true;
            case "svw": unit = HtmlCssLengthUnit.Svw; return true;
            case "svh": unit = HtmlCssLengthUnit.Svh; return true;
            case "svmin": unit = HtmlCssLengthUnit.Svmin; return true;
            case "svmax": unit = HtmlCssLengthUnit.Svmax; return true;
            case "lvw": unit = HtmlCssLengthUnit.Lvw; return true;
            case "lvh": unit = HtmlCssLengthUnit.Lvh; return true;
            case "lvmin": unit = HtmlCssLengthUnit.Lvmin; return true;
            case "lvmax": unit = HtmlCssLengthUnit.Lvmax; return true;
            case "dvw": unit = HtmlCssLengthUnit.Dvw; return true;
            case "dvh": unit = HtmlCssLengthUnit.Dvh; return true;
            case "dvmin": unit = HtmlCssLengthUnit.Dvmin; return true;
            case "dvmax": unit = HtmlCssLengthUnit.Dvmax; return true;
            case "cqw": unit = HtmlCssLengthUnit.Cqw; return true;
            case "cqh": unit = HtmlCssLengthUnit.Cqh; return true;
            case "cqi": unit = HtmlCssLengthUnit.Cqi; return true;
            case "cqb": unit = HtmlCssLengthUnit.Cqb; return true;
            case "cqmin": unit = HtmlCssLengthUnit.Cqmin; return true;
            case "cqmax": unit = HtmlCssLengthUnit.Cqmax; return true;
            default: unit = default; canonical = null; return false;
        }
    }
}
