namespace OfficeIMO.Html.Css;

public static partial class HtmlCssMathResolver {
    /// <summary>Classifies computed math components before a layout-dependent percentage is resolved.</summary>
    internal static bool MixesPercentageAndNonZeroLength(HtmlCssMathExpression expression, HtmlCssLengthResolutionContext context) {
        ComputedComponents components = ResolveComputedComponents(expression, context);
        return expression.IsCalculated && components.HasPercentage && components.HasNonZeroLength;
    }

    private static ComputedComponents ResolveComputedComponents(HtmlCssMathExpression expression, HtmlCssLengthResolutionContext context) {
        if (expression.Type == HtmlCssNumericType.Number) return default;
        if (expression.Type == HtmlCssNumericType.Percentage) return new ComputedComponents(0D, true, false);
        if (expression.Type == HtmlCssNumericType.Length) {
            Evaluation computedLength = Evaluate(expression, context, default);
            return computedLength.IsResolved
                ? new ComputedComponents(computedLength.Value, false, false)
                : new ComputedComponents(0D, false, true);
        }
        if (expression.Kind == HtmlCssMathExpressionKind.Calc) return ResolveComputedComponents(expression.Children[0], context);
        if (expression.Kind is HtmlCssMathExpressionKind.Add or HtmlCssMathExpressionKind.Subtract) {
            ComputedComponents left = ResolveComputedComponents(expression.Children[0], context);
            ComputedComponents right = ResolveComputedComponents(expression.Children[1], context);
            double sign = expression.Kind == HtmlCssMathExpressionKind.Add ? 1D : -1D;
            return new ComputedComponents(left.Length + sign * right.Length,
                left.HasPercentage || right.HasPercentage, left.ComparisonLength || right.ComparisonLength);
        }
        if (expression.Kind is HtmlCssMathExpressionKind.Multiply or HtmlCssMathExpressionKind.Divide) {
            HtmlCssMathExpression left = expression.Children[0];
            HtmlCssMathExpression right = expression.Children[1];
            bool numberOnLeft = left.Type == HtmlCssNumericType.Number;
            HtmlCssMathExpression number = numberOnLeft ? left : right;
            HtmlCssMathExpression dimension = numberOnLeft ? right : left;
            Evaluation scale = Evaluate(number, context, default);
            ComputedComponents components = ResolveComputedComponents(dimension, context);
            if (!scale.IsResolved || expression.Kind == HtmlCssMathExpressionKind.Divide && scale.Value == 0D) return components;
            double multiplier = expression.Kind == HtmlCssMathExpressionKind.Divide ? 1D / scale.Value : scale.Value;
            return new ComputedComponents(components.Length * multiplier, components.HasPercentage,
                components.ComparisonLength && multiplier != 0D);
        }
        bool hasPercentage = false;
        bool hasLength = false;
        foreach (HtmlCssMathExpression child in expression.Children) {
            ComputedComponents components = ResolveComputedComponents(child, context);
            hasPercentage |= components.HasPercentage;
            hasLength |= components.HasNonZeroLength;
        }
        if (!hasPercentage) {
            Evaluation comparison = Evaluate(expression, context, default);
            if (comparison.IsResolved) return new ComputedComponents(comparison.Value, false, false);
        }
        // A min/max/clamp branch can contain a nonzero length even when its
        // value at a zero percentage basis happens to be zero.
        return new ComputedComponents(0D, hasPercentage, hasLength);
    }

    private readonly struct ComputedComponents {
        internal ComputedComponents(double length, bool hasPercentage, bool comparisonLength) {
            Length = length;
            HasPercentage = hasPercentage;
            ComparisonLength = comparisonLength;
        }

        internal double Length { get; }
        internal bool HasPercentage { get; }
        internal bool ComparisonLength { get; }
        internal bool HasNonZeroLength => Length != 0D || ComparisonLength;
    }
}
