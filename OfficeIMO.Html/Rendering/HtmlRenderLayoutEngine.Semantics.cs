using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderFlowBlock ApplyElementSemantics(HtmlRenderFlowBlock block, IElement element) {
        HtmlRenderFlowBlock listBlock = ApplyListSemantics(block, element);
        if (!TryResolveSemanticGroupRole(element.TagName, out HtmlRenderSemanticGroupRole role)) return listBlock;
        return listBlock.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                role,
                0D,
                0D,
                Math.Max(0.01D, listBlock.Width),
                Math.Max(0.01D, listBlock.Height),
                listBlock.Visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        });
    }

    private static bool TryResolveSemanticGroupRole(string tagName, out HtmlRenderSemanticGroupRole role) {
        string tag = tagName.ToLowerInvariant();
        if (tag == "p") {
            role = HtmlRenderSemanticGroupRole.Paragraph;
            return true;
        }
        if (tag == "h1") {
            role = HtmlRenderSemanticGroupRole.Heading1;
            return true;
        }
        if (tag == "h2") {
            role = HtmlRenderSemanticGroupRole.Heading2;
            return true;
        }
        if (tag == "h3") {
            role = HtmlRenderSemanticGroupRole.Heading3;
            return true;
        }
        if (tag == "h4") {
            role = HtmlRenderSemanticGroupRole.Heading4;
            return true;
        }
        if (tag == "h5") {
            role = HtmlRenderSemanticGroupRole.Heading5;
            return true;
        }
        if (tag == "h6") {
            role = HtmlRenderSemanticGroupRole.Heading6;
            return true;
        }
        if (tag == "main" || tag == "section" || tag == "article" || tag == "nav" || tag == "aside") {
            role = HtmlRenderSemanticGroupRole.Section;
            return true;
        }
        if (tag == "header" || tag == "footer") {
            role = HtmlRenderSemanticGroupRole.Division;
            return true;
        }

        role = default;
        return false;
    }
}
