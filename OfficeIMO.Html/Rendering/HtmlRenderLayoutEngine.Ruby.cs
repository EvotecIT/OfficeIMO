using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineRubyRun(
        IElement element,
        double availableWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle rubyStyle,
        string? link,
        double paintOffsetX,
        double paintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        List<INode> baseNodes = element.ChildNodes
            .Where(node => node is not IElement child || !IsRubyAnnotationContainer(child))
            .ToList();
        List<IElement> annotations = element.Children
            .Where(child => child.TagName.Equals("RT", StringComparison.OrdinalIgnoreCase))
            .Concat(element.Children
                .Where(child => child.TagName.Equals("RTC", StringComparison.OrdinalIgnoreCase))
                .SelectMany(child => child.Children)
                .Where(child => child.TagName.Equals("RT", StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (annotations.Count == 0) {
            foreach (INode child in baseNodes) {
                CollectInlineRuns(child, availableWidth, ResolveContainingBlockHeight(parentStyle), rubyStyle, link,
                    depth + 1, paintOffsetX, paintOffsetY, runs);
            }
            return;
        }

        HtmlInlineLayout baseLayout = LayoutInlineNodes(baseNodes, availableWidth, rubyStyle, depth + 1, null, element);
        var annotationVisuals = new List<HtmlRenderVisual>();
        double annotationWidth = 0D;
        double annotationHeight = 0D;
        foreach (IElement annotation in annotations) {
            HtmlRenderBoxStyle annotationStyle = _styleResolver.Resolve(annotation, availableWidth, rubyStyle);
            if (!_styleResolver.IsPropertySpecified(annotation, "font-size")) {
                annotationStyle = annotationStyle.Clone();
                annotationStyle.Font = annotationStyle.Font.WithSize(Math.Max(1D, rubyStyle.Font.Size * 0.5D));
                annotationStyle.LineHeight = Math.Max(annotationStyle.Font.Size, rubyStyle.LineHeight * 0.5D);
            }
            HtmlInlineLayout annotationLayout = LayoutInlineNodes(
                annotation.ChildNodes,
                availableWidth,
                annotationStyle,
                depth + 1,
                null,
                annotation);
            foreach (HtmlRenderVisual visual in annotationLayout.Visuals) {
                annotationVisuals.Add(visual.Translate(annotationWidth, 0D, annotationVisuals.Count));
            }
            annotationWidth += MaximumVerticalSourceRight(annotationLayout.Visuals);
            annotationHeight = Math.Max(annotationHeight, annotationLayout.Height);
        }

        double baseWidth = MaximumVerticalSourceRight(baseLayout.Visuals);
        double rubyWidth = Math.Max(0.01D, Math.Max(baseWidth, annotationWidth));
        double baseOffsetX = ResolveRubyAlignmentOffset(rubyStyle.RubyAlign, rubyWidth, baseWidth);
        double annotationOffsetX = ResolveRubyAlignmentOffset(rubyStyle.RubyAlign, rubyWidth, annotationWidth);
        double baseOffsetY = rubyStyle.RubyPosition == "under" ? 0D : annotationHeight;
        double annotationOffsetY = rubyStyle.RubyPosition == "under" ? baseLayout.Height : 0D;
        var visuals = new List<HtmlRenderVisual>(baseLayout.Visuals.Count + annotationVisuals.Count);
        foreach (HtmlRenderVisual visual in baseLayout.Visuals) {
            visuals.Add(visual.Translate(baseOffsetX, baseOffsetY, visuals.Count));
        }
        foreach (HtmlRenderVisual visual in annotationVisuals) {
            visuals.Add(visual.Translate(annotationOffsetX, annotationOffsetY, visuals.Count));
        }

        double rubyHeight = Math.Max(0.01D, baseLayout.Height + annotationHeight);
        var atomic = new HtmlRenderFlowBlock(
            rubyWidth,
            rubyHeight,
            visuals,
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            HtmlRenderStyleResolver.DescribeSource(element),
            ownerElement: element);
        double baseline = rubyStyle.RubyPosition == "under" ? baseLayout.Height : rubyHeight;
        runs.Add(new HtmlInlineRun(
            atomic,
            rubyStyle,
            link,
            HtmlRenderStyleResolver.DescribeSource(element),
            paintOffsetX,
            paintOffsetY,
            element,
            isReplacedImage: true,
            atomicBaseline: baseline));
    }

    private static bool IsRubyAnnotationContainer(IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        return tag == "rt" || tag == "rp" || tag == "rtc";
    }

    private static double ResolveRubyAlignmentOffset(string rubyAlign, double containerWidth, double contentWidth) {
        double remaining = Math.Max(0D, containerWidth - contentWidth);
        return rubyAlign == "start" ? 0D : remaining / 2D;
    }
}
