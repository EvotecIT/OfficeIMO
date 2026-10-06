using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        private LayoutBox? FontRadical(OfficeMathExpression expression, double scale) {
            if (_mathConstants == null || !TryGlyph("√", scale, out _, out var data, out int glyph)
                || !data!.Vertical.ContainsKey(glyph)) return null;
            LayoutBox content = CrampedLayout(expression.Children[0], scale);
            double thickness = Math.Max(0D, MathValue(OfficeMathConstant.RadicalRuleThickness, scale));
            double gap = Math.Max(0D, MathValue(_compact ? OfficeMathConstant.RadicalVerticalGap
                : OfficeMathConstant.RadicalDisplayStyleVerticalGap, scale));
            LayoutBox? radical = StretchGlyph("√", scale, content.Height + gap + thickness);
            if (radical == null) return null;
            double outside = Math.Max(0D, MathValue(OfficeMathConstant.RadicalExtraAscender, scale));
            double contentY = outside + Math.Max(gap + thickness, radical.Height - content.Height);
            double radicalY = contentY + content.Height - radical.Height;
            double radicalX = 0D, indexX = 0D, indexY = 0D;
            LayoutBox? index = null;
            if (expression.Children.Count == 2) {
                // A radical degree is two math depths smaller than its parent.
                double degreeScale = _scriptLevel == 0
                    ? scale * _mathConstants.GetValue(OfficeMathConstant.ScriptScriptPercentScaleDown) / 100D
                    : _scriptLevel == 1 ? ScriptScale(scale) * .71D : scale * .71D * .71D;
                if (!_childEngines.TryGetValue(8, out LayoutEngine? degreeEngine)) {
                    degreeEngine = new LayoutEngine(_options, true, _cancellationToken, _measureScopedText,
                        _scriptLevel + 2, _cramped);
                    _childEngines.Add(8, degreeEngine);
                }
                index = degreeEngine.Layout(expression.Children[1], degreeScale);
                double before = MathValue(OfficeMathConstant.RadicalKernBeforeDegree, scale);
                double after = MathValue(OfficeMathConstant.RadicalKernAfterDegree, scale);
                radicalX = Math.Max(0D, before + index.Width + after);
                indexX = radicalX - index.Width - after;
                double raise = _mathConstants.GetValue(OfficeMathConstant.RadicalDegreeBottomRaisePercent) / 100D;
                indexY = radicalY + radical.Height * (1D - raise) - index.Height;
            }
            double left = Math.Min(0D, indexX), top = Math.Min(0D, indexY);
            double contentX = radicalX + radical.Width;
            double width = Math.Max(contentX + content.Width, index == null ? 0D : indexX + index.Width) - left;
            double bottom = Math.Max(contentY + content.Height, index == null ? 0D : indexY + index.Height);
            var box = new LayoutBox(width, bottom - top, contentY + content.Baseline - top);
            box.Add(radical, radicalX - left, radicalY - top);
            box.Add(content, contentX - left, contentY - top);
            if (index != null) box.Add(index, indexX - left, indexY - top);
            double ruleY = radicalY + thickness / 2D - top;
            box.Commands.Add(LayoutCommand.Line(contentX - left,
                ruleY, width, ruleY, thickness));
            return box;
        }

        private LayoutBox? FontDelimited(LayoutBox content, string leftText, string rightText, double scale, double gap) {
            if (_mathConstants == null) return null;
            double axis = MathAxis(scale);
            double target = 2D * Math.Max(content.Baseline - axis, content.Height - content.Baseline + axis);
            LayoutBox? left = leftText.Length == 0 ? Text(string.Empty, scale) : StretchGlyph(leftText, scale, target);
            LayoutBox? right = rightText.Length == 0 ? Text(string.Empty, scale) : StretchGlyph(rightText, scale, target);
            if (left == null || right == null) return null;
            double axisY = content.Baseline - axis;
            double leftY = axisY - left.Height / 2D, rightY = axisY - right.Height / 2D;
            double top = Math.Min(0D, Math.Min(leftY, rightY));
            double bottom = Math.Max(content.Height, Math.Max(leftY + left.Height, rightY + right.Height));
            var box = new LayoutBox(left.Width + gap + content.Width + gap + right.Width, bottom - top, content.Baseline - top);
            box.Add(left, 0D, leftY - top); box.Add(content, left.Width + gap, -top);
            box.Add(right, left.Width + gap + content.Width + gap, rightY - top);
            return box;
        }

        private LayoutBox LargeOperator(string text, double scale) {
            LayoutBox box = _mathConstants == null || _compact ? Text(text, scale)
                : StretchGlyph(text, scale, Math.Max(0D, MathValue(OfficeMathConstant.DisplayOperatorMinHeight, scale)),
                    useAssembly: false) ?? Text(text, scale * 1.35D);
            box.LargeOperator = true;
            return box;
        }
    }
}
