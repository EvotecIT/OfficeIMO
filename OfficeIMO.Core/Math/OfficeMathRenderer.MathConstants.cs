using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        private double MathValue(OfficeMathConstant constant, double scale) =>
            _mathConstants!.GetValue(constant) * FontSize(scale) / _mathConstants.UnitsPerEm;

        private double ScriptScale(double scale) {
            if (_mathConstants == null) return scale * _options.ScriptScale;
            double first = _mathConstants.GetValue(OfficeMathConstant.ScriptPercentScaleDown);
            double second = _mathConstants.GetValue(OfficeMathConstant.ScriptScriptPercentScaleDown);
            // MathML math-depth uses font percentages for the first two depths, then 0.71 per depth.
            return scale * (_scriptLevel == 0 ? first / 100D : _scriptLevel == 1 ? second / first : .71D);
        }

        private LayoutBox FontFraction(LayoutBox numerator, LayoutBox denominator, double scale) {
            double thickness = Math.Max(0D, MathValue(OfficeMathConstant.FractionRuleThickness, scale));
            double axis = MathAxis(scale);
            double numeratorShift = Math.Max(MathValue(_compact ? OfficeMathConstant.FractionNumeratorShiftUp
                    : OfficeMathConstant.FractionNumeratorDisplayStyleShiftUp, scale),
                axis + thickness / 2D + MathValue(_compact ? OfficeMathConstant.FractionNumeratorGapMin
                    : OfficeMathConstant.FractionNumDisplayStyleGapMin, scale) + numerator.Height - numerator.Baseline);
            double denominatorShift = Math.Max(MathValue(_compact ? OfficeMathConstant.FractionDenominatorShiftDown
                    : OfficeMathConstant.FractionDenominatorDisplayStyleShiftDown, scale),
                thickness / 2D + MathValue(_compact ? OfficeMathConstant.FractionDenominatorGapMin
                    : OfficeMathConstant.FractionDenomDisplayStyleGapMin, scale) + denominator.Baseline - axis);
            double baseline = Math.Max(0D, Math.Max(numeratorShift + numerator.Baseline, axis + thickness / 2D));
            double descent = Math.Max(0D, Math.Max(denominatorShift + denominator.Height - denominator.Baseline,
                thickness / 2D - axis));
            double inset = Math.Max(FontSize(scale) / 20D, thickness / 2D);
            double width = Math.Max(numerator.Width, denominator.Width) + inset * 2D;
            var box = new LayoutBox(width, baseline + descent, baseline);
            box.Add(numerator, (width - numerator.Width) / 2D, baseline - (numeratorShift + numerator.Baseline));
            box.Commands.Add(LayoutCommand.Line(inset / 2D, baseline - axis, width - inset / 2D, baseline - axis, thickness));
            box.Add(denominator, (width - denominator.Width) / 2D, baseline + denominatorShift - denominator.Baseline);
            return box;
        }

        private LayoutBox FontScripts(LayoutBox basis, LayoutBox? sub, LayoutBox? sup, double scale, bool left) {
            double subShift = sub == null ? 0D : Math.Max(MathValue(OfficeMathConstant.SubscriptShiftDown, scale),
                Math.Max(sub.Baseline - MathValue(OfficeMathConstant.SubscriptTopMax, scale),
                    MathValue(OfficeMathConstant.SubscriptBaselineDropMin, scale) + basis.Height - basis.Baseline));
            double superShift = sup == null ? 0D : Math.Max(MathValue(_cramped
                    ? OfficeMathConstant.SuperscriptShiftUpCramped : OfficeMathConstant.SuperscriptShiftUp, scale),
                Math.Max(MathValue(OfficeMathConstant.SuperscriptBottomMin, scale) + sup.Height - sup.Baseline,
                    basis.Baseline - MathValue(OfficeMathConstant.SuperscriptBaselineDropMax, scale)));
            if (sub != null && sup != null) {
                double gap = subShift - sub.Baseline + superShift - (sup.Height - sup.Baseline);
                double missing = MathValue(OfficeMathConstant.SubSuperscriptGapMin, scale) - gap;
                if (missing > 0D) {
                    double superAdjustment = Math.Min(missing, Math.Max(0D,
                        MathValue(OfficeMathConstant.SuperscriptBottomMaxWithSubscript, scale)
                        - superShift + sup.Height - sup.Baseline));
                    superShift += superAdjustment;
                    subShift += missing - superAdjustment;
                }
            }
            double baseline = Math.Max(basis.Baseline, Math.Max(
                sub == null ? 0D : sub.Baseline - subShift, sup == null ? 0D : sup.Baseline + superShift));
            double descent = Math.Max(basis.Height - basis.Baseline, Math.Max(
                sub == null ? 0D : sub.Height - sub.Baseline + subShift,
                sup == null ? 0D : sup.Height - sup.Baseline - superShift));
            double space = Math.Max(0D, MathValue(OfficeMathConstant.SpaceAfterScript, scale));
            double italic = basis.LargeOperator ? 0D : basis.ItalicCorrection;
            double largeItalic = basis.LargeOperator ? basis.ItalicCorrection : 0D;
            double supKern = sup == null ? 0D : ScriptKern(basis, sup, superShift, over: true, left);
            double subKern = sub == null ? 0D : ScriptKern(basis, sub, -subShift, over: false, left);
            double baseOrigin = basis.GlyphUnit > 0D ? basis.GlyphOrigin : 0D;
            double baseAdvance = basis.GlyphUnit > 0D ? basis.GlyphAdvance : basis.Width;
            double ScriptX(LayoutBox? script, double correction, double kern) {
                if (script == null) return 0D;
                double origin = script.GlyphUnit > 0D ? script.GlyphOrigin : 0D;
                double advance = script.GlyphUnit > 0D ? script.GlyphAdvance : script.Width;
                return left ? baseOrigin - advance - origin - correction - kern
                    : baseOrigin + baseAdvance + correction + kern - origin;
            }
            double supX = ScriptX(sup, italic, supKern);
            double subX = ScriptX(sub, -largeItalic, subKern);
            double start = Math.Min(0D, Math.Min(sup == null ? 0D : supX, sub == null ? 0D : subX));
            double end = Math.Max(basis.Width, Math.Max(sup == null ? 0D : supX + sup.Width, sub == null ? 0D : subX + sub.Width));
            double origin = -start + (left ? space : 0D);
            var box = new LayoutBox(end - start + space, baseline + descent, baseline);
            box.Add(basis, origin, baseline - basis.Baseline);
            if (sup != null) box.Add(sup, origin + supX, baseline - (superShift + sup.Baseline));
            if (sub != null) box.Add(sub, origin + subX, baseline + subShift - sub.Baseline);
            return box;
        }

        private LayoutBox FontLimit(LayoutBox basis, LayoutBox limit, double scale, bool over) {
            double gap = over
                ? Math.Max(MathValue(OfficeMathConstant.UpperLimitGapMin, scale),
                    MathValue(OfficeMathConstant.UpperLimitBaselineRiseMin, scale) - (limit.Height - limit.Baseline))
                : Math.Max(MathValue(OfficeMathConstant.LowerLimitGapMin, scale),
                    MathValue(OfficeMathConstant.LowerLimitBaselineDropMin, scale) - limit.Baseline);
            gap = Math.Max(0D, gap);
            double center = basis.LimitCenter ?? basis.Width / 2D;
            double limitX = center - limit.Width / 2D + (over ? 1D : -1D) * basis.ItalicCorrection / 2D;
            double left = Math.Min(0D, limitX);
            double width = Math.Max(basis.Width, limitX + limit.Width) - left;
            double basisY = over ? limit.Height + gap : 0D;
            double limitY = over ? 0D : basis.Height + gap;
            var box = new LayoutBox(width, basis.Height + gap + limit.Height, basisY + basis.Baseline);
            box.Add(basis, -left, basisY);
            box.Add(limit, limitX - left, limitY);
            box.ItalicCorrection = basis.ItalicCorrection;
            box.LimitCenter = center - left;
            return box;
        }

        private LayoutBox FontBar(LayoutBox content, double scale, bool over) {
            double thickness = Math.Max(0D, MathValue(over ? OfficeMathConstant.OverbarRuleThickness
                : OfficeMathConstant.UnderbarRuleThickness, scale));
            double gap = Math.Max(0D, MathValue(over ? OfficeMathConstant.OverbarVerticalGap
                : OfficeMathConstant.UnderbarVerticalGap, scale));
            double outside = Math.Max(0D, MathValue(over ? OfficeMathConstant.OverbarExtraAscender
                : OfficeMathConstant.UnderbarExtraDescender, scale));
            double extra = outside + gap + thickness;
            var box = new LayoutBox(content.Width, content.Height + extra, content.Baseline + (over ? extra : 0D));
            box.Add(content, 0D, over ? extra : 0D);
            double y = over ? outside + thickness / 2D : content.Height + gap + thickness / 2D;
            box.Commands.Add(LayoutCommand.Line(0D, y, content.Width, y, thickness));
            return box;
        }
    }
}
