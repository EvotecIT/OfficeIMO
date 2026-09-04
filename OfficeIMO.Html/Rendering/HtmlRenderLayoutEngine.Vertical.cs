using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool IsVerticalWritingMode(string writingMode) =>
        writingMode == "vertical-rl" || writingMode == "vertical-lr"
        || writingMode == "sideways-rl" || writingMode == "sideways-lr";

    private double ResolveVerticalInlineExtent(
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        double fallback) {
        if (style.ExplicitHeight.HasValue) {
            double height = style.ExplicitHeight.Value;
            return Math.Max(1D, style.BorderBox ? height - style.VerticalInsets : height);
        }

        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (containingHeight.HasValue) return Math.Max(1D, containingHeight.Value - style.VerticalInsets);
        double surfaceHeight = _options.Mode == HtmlRenderMode.Paged
            ? _activePageGeometry.ContentHeight
            : (_options.ViewportHeight ?? fallback) - _options.Margins.Top - _options.Margins.Bottom;
        return Math.Max(1D, surfaceHeight - style.VerticalInsets);
    }

    private HtmlInlineLayout TransformSidewaysVerticalInlineLayout(
        HtmlInlineLayout inline,
        HtmlRenderBoxStyle style,
        IElement element) {
        if (inline.Visuals.Count == 0) return inline;
        double sourceWidth = Math.Max(0.01D, MaximumVerticalSourceRight(inline.Visuals));
        double sourceHeight = Math.Max(0.01D, inline.Height);
        bool rightToLeftBlocks = style.WritingMode == "vertical-rl" || style.WritingMode == "sideways-rl";
        OfficeTransform transform = rightToLeftBlocks
            ? OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(sourceHeight, 0D))
            : OfficeTransform.RotateDegrees(-90D).Then(OfficeTransform.Translate(0D, sourceWidth));
        string source = HtmlRenderStyleResolver.DescribeSource(element) + ":vertical-writing";
        bool sidewaysOnly = style.TextOrientation == "sideways"
            || style.WritingMode.StartsWith("sideways", StringComparison.Ordinal);
        IReadOnlyList<HtmlRenderVisual> verticalVisuals = sidewaysOnly
            ? new HtmlRenderVisual[] {
                new HtmlRenderEffectGroup(
                    0D,
                    0D,
                    sourceHeight,
                    sourceWidth,
                    transform,
                    1D,
                    inline.Visuals,
                    0,
                    source)
            }
            : TransformVerticalVisuals(
                inline.Visuals,
                transform,
                style.TextOrientation,
                rightToLeftBlocks,
                sourceHeight,
                sourceWidth,
                source);
        string logicalText = ResolveLogicalText(inline.Visuals, string.Empty);
        HtmlRenderVisual verticalVisual = logicalText.Length == 0
            ? new HtmlRenderEffectGroup(
                0D,
                0D,
                sourceHeight,
                sourceWidth,
                OfficeTransform.Identity,
                1D,
                verticalVisuals,
                0,
                source)
            : new HtmlRenderLogicalTextGroup(
                logicalText,
                0D,
                0D,
                sourceHeight,
                sourceWidth,
                verticalVisuals,
                0,
                source);
        return new HtmlInlineLayout(
            new[] { verticalVisual },
            sourceWidth,
            runningStringAssignments: inline.RunningStringAssignments);
    }

    private IReadOnlyList<HtmlRenderVisual> TransformVerticalVisuals(
        IReadOnlyList<HtmlRenderVisual> visuals,
        OfficeTransform axisTransform,
        string textOrientation,
        bool rightToLeftBlocks,
        double destinationWidth,
        double destinationHeight,
        string source) {
        var result = new List<HtmlRenderVisual>();
        for (int index = 0; index < visuals.Count; index++) {
            HtmlRenderVisual visual = visuals[index];
            if (visual is HtmlRenderText text) {
                AppendVerticalText(result, text, axisTransform, textOrientation, rightToLeftBlocks, source);
                continue;
            }
            if (visual is HtmlRenderSemanticGroup semantic) {
                IReadOnlyList<HtmlRenderVisual> children = TransformVerticalVisuals(
                    semantic.Visuals,
                    axisTransform,
                    textOrientation,
                    rightToLeftBlocks,
                    destinationWidth,
                    destinationHeight,
                    source);
                result.Add(new HtmlRenderSemanticGroup(
                    semantic.Role,
                    0D,
                    0D,
                    destinationWidth,
                    destinationHeight,
                    children,
                    result.Count,
                    semantic.Source,
                    semantic.ColumnSpan,
                    semantic.RowSpan,
                    semantic.HeaderScope,
                    structureElementKey: semantic.StructureElementKey));
                continue;
            }
            if (visual is HtmlRenderLogicalTextGroup logical) {
                IReadOnlyList<HtmlRenderVisual> children = TransformVerticalVisuals(
                    logical.Visuals,
                    axisTransform,
                    textOrientation,
                    rightToLeftBlocks,
                    destinationWidth,
                    destinationHeight,
                    source);
                result.Add(new HtmlRenderLogicalTextGroup(
                    logical.Text,
                    0D,
                    0D,
                    destinationWidth,
                    destinationHeight,
                    children,
                    result.Count,
                    logical.Source));
                continue;
            }

            result.Add(new HtmlRenderEffectGroup(
                0D,
                0D,
                destinationWidth,
                destinationHeight,
                axisTransform,
                1D,
                new[] { visual },
                result.Count,
                source));
        }
        return result;
    }

    private void AppendVerticalText(
        List<HtmlRenderVisual> destination,
        HtmlRenderText visual,
        OfficeTransform axisTransform,
        string textOrientation,
        bool rightToLeftBlocks,
        string source) {
        IReadOnlyList<string> elements = OfficeTextElements.Split(visual.Text);
        if (elements.Count == 0) return;

        var measured = new double[elements.Count];
        double measuredTotal = 0D;
        for (int index = 0; index < elements.Count; index++) {
            measured[index] = Math.Max(0.01D, MeasureText(elements[index], visual.Font));
            measuredTotal += measured[index];
        }

        double totalAdvance = Math.Abs(visual.TextAdvanceWidth ?? visual.Width);
        double scale = measuredTotal > 0D ? totalAdvance / measuredTotal : 1D;
        double cursor = visual.X;
        for (int index = 0; index < elements.Count; index++) {
            string textElement = elements[index];
            double glyphWidth = measured[index];
            double advance = measured[index] * scale;
            OfficePoint center = axisTransform.TransformPoint(new OfficePoint(
                cursor + advance / 2D,
                visual.Y + visual.Height / 2D));
            cursor += advance;

            bool upright = textOrientation == "upright" || IsMixedOrientationUpright(textElement);
            double x = center.X - glyphWidth / 2D;
            double y = center.Y - visual.Height / 2D;
            var glyph = new HtmlRenderText(
                textElement,
                x,
                y,
                glyphWidth,
                visual.Height,
                visual.Font,
                visual.Color,
                visual.Alignment,
                visual.LineHeight,
                0,
                visual.LinkUri,
                visual.Source,
                visual.SemanticRole,
                layoutY: y,
                visual.SemanticNodeId,
                glyphWidth,
                visual.BidiVisualOrderResolved,
                visual.SemanticFragmentOrder,
                visual.LogicalTextOrder,
                visual.UnderlineStyle,
                visual.StrikethroughStyle,
                visual.Baseline,
                visual.BaselineLevel,
                visual.BaselineScale,
                visual.BaselineOffset,
                glyphWidth,
                visual.DecorationColor);
            if (upright) {
                destination.Add(glyph.Translate(0D, 0D, destination.Count));
                continue;
            }

            destination.Add(new HtmlRenderEffectGroup(
                center.X - visual.Height / 2D,
                center.Y - glyphWidth / 2D,
                visual.Height,
                glyphWidth,
                OfficeTransform.RotateDegrees(rightToLeftBlocks ? 90D : -90D, center.X, center.Y),
                1D,
                new[] { glyph },
                destination.Count,
                source));
        }
    }

    private static bool IsMixedOrientationUpright(string textElement) {
        if (string.IsNullOrEmpty(textElement) || char.IsWhiteSpace(textElement, 0)) return true;
        int codePoint = char.ConvertToUtf32(textElement, 0);
        return (codePoint >= 0x1100 && codePoint <= 0x11FF)
            || (codePoint >= 0x2E80 && codePoint <= 0xA4CF)
            || (codePoint >= 0xAC00 && codePoint <= 0xD7AF)
            || (codePoint >= 0xF900 && codePoint <= 0xFAFF)
            || (codePoint >= 0xFE10 && codePoint <= 0xFE1F)
            || (codePoint >= 0xFE30 && codePoint <= 0xFE6F)
            || (codePoint >= 0xFF01 && codePoint <= 0xFF60)
            || (codePoint >= 0xFFE0 && codePoint <= 0xFFE6)
            || codePoint >= 0x1F000;
    }

    private static double MaximumVerticalSourceRight(IEnumerable<HtmlRenderVisual> visuals) {
        double maximum = 0.01D;
        foreach (HtmlRenderVisual visual in visuals) {
            maximum = Math.Max(maximum, visual.X + visual.Width);
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children != null) maximum = Math.Max(maximum, MaximumVerticalSourceRight(children));
        }
        return maximum;
    }
}
