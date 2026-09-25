using System;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Resolves owned CSS length expressions when the consuming layout context is available.</summary>
public static class HtmlCssMathResolver {
    private const double CssPixelsPerInch = 96D;

    /// <summary>
    /// Resolves a length, percentage, or length-percentage expression to CSS pixels.
    /// Missing context is returned as data so callers can preserve their normal fallback behavior.
    /// </summary>
    public static HtmlCssLengthResolutionResult ResolveLength(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context,
        CancellationToken cancellationToken = default) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        if (context == null) throw new ArgumentNullException(nameof(context));
        cancellationToken.ThrowIfCancellationRequested();
        if (expression.Type != HtmlCssNumericType.Length
            && expression.Type != HtmlCssNumericType.Percentage
            && expression.Type != HtmlCssNumericType.LengthPercentage)
            return Result(HtmlCssLengthResolutionStatus.IncompatibleType);
        Evaluation evaluated = Evaluate(expression, context, cancellationToken);
        if (evaluated.Status != HtmlCssLengthResolutionStatus.Resolved) return Result(evaluated.Status);
        if (!IsFinite(evaluated.Value)) return Result(HtmlCssLengthResolutionStatus.NonFiniteValue);
        return new HtmlCssLengthResolutionResult(HtmlCssLengthResolutionStatus.Resolved, evaluated.Value);
    }

    private static Evaluation Evaluate(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (expression.Kind) {
            case HtmlCssMathExpressionKind.Literal:
                return EvaluateLiteral(expression, context);
            case HtmlCssMathExpressionKind.Add:
            case HtmlCssMathExpressionKind.Subtract:
            case HtmlCssMathExpressionKind.Multiply:
            case HtmlCssMathExpressionKind.Divide:
                return EvaluateBinary(expression, context, cancellationToken);
            case HtmlCssMathExpressionKind.Calc:
                return EvaluateGrouping(expression, context, cancellationToken);
            case HtmlCssMathExpressionKind.Min:
            case HtmlCssMathExpressionKind.Max:
            case HtmlCssMathExpressionKind.Clamp:
                return EvaluateComparison(expression, context, cancellationToken);
            default:
                return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        }
    }

    private static Evaluation EvaluateLiteral(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context) {
        if (!expression.Value.HasValue || !IsFinite(expression.Value.Value))
            return Evaluation.Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
        double value = expression.Value.Value;
        if (expression.Type == HtmlCssNumericType.Number) return Evaluation.Success(value);
        if (expression.Type == HtmlCssNumericType.Percentage) {
            Evaluation reference = ContextValue(context.PercentageReference,
                HtmlCssLengthResolutionStatus.MissingPercentageReference);
            return reference.Scale(value / 100D);
        }
        if (expression.Type != HtmlCssNumericType.Length)
            return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        if (!expression.Unit.HasValue) return value == 0D
            ? Evaluation.Success(0D)
            : Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        switch (expression.Unit.Value) {
            case HtmlCssLengthUnit.Px: return Evaluation.Success(value);
            case HtmlCssLengthUnit.Pt: return Evaluation.Success(value * CssPixelsPerInch / 72D);
            case HtmlCssLengthUnit.Pc: return Evaluation.Success(value * CssPixelsPerInch / 6D);
            case HtmlCssLengthUnit.In: return Evaluation.Success(value * CssPixelsPerInch);
            case HtmlCssLengthUnit.Cm: return Evaluation.Success(value * CssPixelsPerInch / 2.54D);
            case HtmlCssLengthUnit.Mm: return Evaluation.Success(value * CssPixelsPerInch / 25.4D);
            case HtmlCssLengthUnit.Q: return Evaluation.Success(value * CssPixelsPerInch / 101.6D);
        }
        Evaluation scale = ResolveUnitScale(expression.Unit.Value, context);
        return scale.Scale(value);
    }

    private static Evaluation EvaluateBinary(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context,
        CancellationToken cancellationToken) {
        if (expression.Children.Count != 2)
            return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        Evaluation left = Evaluate(expression.Children[0], context, cancellationToken);
        if (!left.IsResolved) return left;
        Evaluation right = Evaluate(expression.Children[1], context, cancellationToken);
        if (!right.IsResolved) return right;
        double value;
        switch (expression.Kind) {
            case HtmlCssMathExpressionKind.Add: value = left.Value + right.Value; break;
            case HtmlCssMathExpressionKind.Subtract: value = left.Value - right.Value; break;
            case HtmlCssMathExpressionKind.Multiply: value = left.Value * right.Value; break;
            case HtmlCssMathExpressionKind.Divide:
                if (right.Value == 0D) return Evaluation.Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
                value = left.Value / right.Value;
                break;
            default: return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        }
        return IsFinite(value) ? Evaluation.Success(value) : Evaluation.Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
    }

    private static Evaluation EvaluateGrouping(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context,
        CancellationToken cancellationToken) {
        if (expression.Children.Count != 1)
            return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        return Evaluate(expression.Children[0], context, cancellationToken);
    }

    private static Evaluation EvaluateComparison(
        HtmlCssMathExpression expression,
        HtmlCssLengthResolutionContext context,
        CancellationToken cancellationToken) {
        int required = expression.Kind == HtmlCssMathExpressionKind.Clamp ? 3 : 1;
        if (expression.Children.Count < required
            || expression.Kind == HtmlCssMathExpressionKind.Clamp && expression.Children.Count != 3)
            return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        var values = new double[expression.Children.Count];
        for (int index = 0; index < expression.Children.Count; index++) {
            Evaluation value = Evaluate(expression.Children[index], context, cancellationToken);
            if (!value.IsResolved) return value;
            values[index] = value.Value;
        }
        double resolved = values[0];
        if (expression.Kind == HtmlCssMathExpressionKind.Min) {
            for (int index = 1; index < values.Length; index++) resolved = Math.Min(resolved, values[index]);
        } else if (expression.Kind == HtmlCssMathExpressionKind.Max) {
            for (int index = 1; index < values.Length; index++) resolved = Math.Max(resolved, values[index]);
        } else {
            resolved = Math.Max(values[0], Math.Min(values[1], values[2]));
        }
        return IsFinite(resolved) ? Evaluation.Success(resolved) : Evaluation.Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
    }

    private static Evaluation ResolveUnitScale(HtmlCssLengthUnit unit, HtmlCssLengthResolutionContext context) {
        switch (unit) {
            case HtmlCssLengthUnit.Em:
                return ContextValue(context.FontSize, HtmlCssLengthResolutionStatus.MissingFontSize);
            case HtmlCssLengthUnit.Ex:
                return ContextValue(context.FontSize, HtmlCssLengthResolutionStatus.MissingFontSize).Scale(0.5D);
            case HtmlCssLengthUnit.Rem:
                return ContextValue(context.RootFontSize, HtmlCssLengthResolutionStatus.MissingRootFontSize);
            case HtmlCssLengthUnit.Vw:
                return Viewport(context.ViewportWidth, HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Vh:
                return Viewport(context.ViewportHeight, HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Vmin:
                return ViewportMinMax(context.ViewportWidth, context.ViewportHeight, minimum: true);
            case HtmlCssLengthUnit.Vmax:
                return ViewportMinMax(context.ViewportWidth, context.ViewportHeight, minimum: false);
            case HtmlCssLengthUnit.Svw:
                return Viewport(context.SmallViewportWidth ?? context.ViewportWidth, HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Svh:
                return Viewport(context.SmallViewportHeight ?? context.ViewportHeight, HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Svmin:
                return ViewportMinMax(context.SmallViewportWidth ?? context.ViewportWidth,
                    context.SmallViewportHeight ?? context.ViewportHeight, minimum: true);
            case HtmlCssLengthUnit.Svmax:
                return ViewportMinMax(context.SmallViewportWidth ?? context.ViewportWidth,
                    context.SmallViewportHeight ?? context.ViewportHeight, minimum: false);
            case HtmlCssLengthUnit.Lvw:
                return Viewport(context.LargeViewportWidth ?? context.ViewportWidth, HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Lvh:
                return Viewport(context.LargeViewportHeight ?? context.ViewportHeight, HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Lvmin:
                return ViewportMinMax(context.LargeViewportWidth ?? context.ViewportWidth,
                    context.LargeViewportHeight ?? context.ViewportHeight, minimum: true);
            case HtmlCssLengthUnit.Lvmax:
                return ViewportMinMax(context.LargeViewportWidth ?? context.ViewportWidth,
                    context.LargeViewportHeight ?? context.ViewportHeight, minimum: false);
            case HtmlCssLengthUnit.Dvw:
                return Viewport(context.DynamicViewportWidth ?? context.ViewportWidth, HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Dvh:
                return Viewport(context.DynamicViewportHeight ?? context.ViewportHeight, HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Dvmin:
                return ViewportMinMax(context.DynamicViewportWidth ?? context.ViewportWidth,
                    context.DynamicViewportHeight ?? context.ViewportHeight, minimum: true);
            case HtmlCssLengthUnit.Dvmax:
                return ViewportMinMax(context.DynamicViewportWidth ?? context.ViewportWidth,
                    context.DynamicViewportHeight ?? context.ViewportHeight, minimum: false);
            case HtmlCssLengthUnit.Cqw:
                return Container(context.ContainerWidth, context.SmallViewportWidth ?? context.ViewportWidth,
                    HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Cqh:
                return Container(context.ContainerHeight, context.SmallViewportHeight ?? context.ViewportHeight,
                    HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Cqi:
                return Container(context.ContainerInlineSize ?? context.ContainerWidth,
                    context.SmallViewportWidth ?? context.ViewportWidth, HtmlCssLengthResolutionStatus.MissingViewportWidth);
            case HtmlCssLengthUnit.Cqb:
                return Container(context.ContainerBlockSize ?? context.ContainerHeight,
                    context.SmallViewportHeight ?? context.ViewportHeight, HtmlCssLengthResolutionStatus.MissingViewportHeight);
            case HtmlCssLengthUnit.Cqmin:
                return ContainerMinMax(context, minimum: true);
            case HtmlCssLengthUnit.Cqmax:
                return ContainerMinMax(context, minimum: false);
            default:
                return Evaluation.Fail(HtmlCssLengthResolutionStatus.IncompatibleType);
        }
    }

    private static Evaluation Viewport(double? value, HtmlCssLengthResolutionStatus missing) =>
        ContextValue(value, missing).Scale(0.01D);

    private static Evaluation ViewportMinMax(double? width, double? height, bool minimum) {
        Evaluation x = ContextValue(width, HtmlCssLengthResolutionStatus.MissingViewportWidth);
        if (!x.IsResolved) return x;
        Evaluation y = ContextValue(height, HtmlCssLengthResolutionStatus.MissingViewportHeight);
        if (!y.IsResolved) return y;
        return Evaluation.Success((minimum ? Math.Min(x.Value, y.Value) : Math.Max(x.Value, y.Value)) / 100D);
    }

    private static Evaluation Container(double? container, double? smallViewport, HtmlCssLengthResolutionStatus missing) =>
        ContextValue(container ?? smallViewport, missing).Scale(0.01D);

    private static Evaluation ContainerMinMax(HtmlCssLengthResolutionContext context, bool minimum) {
        Evaluation width = ContextValue(context.ContainerInlineSize ?? context.ContainerWidth
                ?? context.SmallViewportWidth ?? context.ViewportWidth,
            HtmlCssLengthResolutionStatus.MissingViewportWidth);
        if (!width.IsResolved) return width;
        Evaluation height = ContextValue(context.ContainerBlockSize ?? context.ContainerHeight
                ?? context.SmallViewportHeight ?? context.ViewportHeight,
            HtmlCssLengthResolutionStatus.MissingViewportHeight);
        if (!height.IsResolved) return height;
        return Evaluation.Success((minimum ? Math.Min(width.Value, height.Value) : Math.Max(width.Value, height.Value)) / 100D);
    }

    private static Evaluation ContextValue(double? value, HtmlCssLengthResolutionStatus missing) {
        if (!value.HasValue) return Evaluation.Fail(missing);
        return IsFinite(value.Value)
            ? Evaluation.Success(value.Value)
            : Evaluation.Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
    }

    private static HtmlCssLengthResolutionResult Result(HtmlCssLengthResolutionStatus status) =>
        new HtmlCssLengthResolutionResult(status, null);

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private readonly struct Evaluation {
        private Evaluation(HtmlCssLengthResolutionStatus status, double value) { Status = status; Value = value; }
        internal HtmlCssLengthResolutionStatus Status { get; }
        internal double Value { get; }
        internal bool IsResolved => Status == HtmlCssLengthResolutionStatus.Resolved;
        internal Evaluation Scale(double factor) {
            if (!IsResolved) return this;
            double scaled = Value * factor;
            return IsFinite(scaled) ? Success(scaled) : Fail(HtmlCssLengthResolutionStatus.NonFiniteValue);
        }
        internal static Evaluation Success(double value) => new Evaluation(HtmlCssLengthResolutionStatus.Resolved, value);
        internal static Evaluation Fail(HtmlCssLengthResolutionStatus status) => new Evaluation(status, 0D);
    }
}
