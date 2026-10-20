using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void CollectInlineRuns(
        INode node,
        double width,
        double? containingHeight,
        HtmlRenderBoxStyle inheritedStyle,
        string? inheritedLink,
        int depth,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        IList<HtmlInlineRun> runs, bool flattenInternalTableBoxes = false) {
        CheckCancellation();
        ChargeLayoutOperation("inline node traversal");
        if (IsClosedDisclosureChild(node)) return;
        if (depth > _options.MaxLayoutDepth) {
            if (node is IElement limitedElement) EnsureDepth(depth, limitedElement);
            throw new InvalidOperationException("HTML inline layout exceeded the configured maximum depth.");
        }

        if (node is IText textNode) {
            if (textNode.Data.Length > 0) {
                bool contextOnly = inheritedStyle.Font.Size <= 0D;
                if (!contextOnly) ReportUnsupportedComplexTextShaping(textNode.Data, textNode.ParentElement, inheritedStyle);
                string source = textNode.ParentElement != null
                    && HtmlRenderSourceIdentity.TryGet(textNode.ParentElement, out string interactionSource)
                    ? interactionSource
                    : inheritedStyle.SemanticRole;
                runs.Add(new HtmlInlineRun(
                    textNode.Data,
                    inheritedStyle,
                    inheritedLink,
                    source,
                    inheritedPaintOffsetX,
                    inheritedPaintOffsetY,
                    textNode.ParentElement,
                    textTransformPending: true) { IsTextTransformContextOnly = contextOnly, SemanticSourceOrder = GetSemanticSourceOrder(textNode) });
            }

            return;
        }

        if (!(node is IElement element) || ShouldSkipElement(element)) return;
        string tag = element.TagName.ToLowerInvariant();
        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, width, inheritedStyle);
        style = ForwardContainingHeightBasis(element, style, inheritedStyle);
        style = ResolveColumnFragmentStyle(element, PrepareButtonChildStyle(element, style));
        if (style.FloatSide == "footnote" && _options.Mode != HtmlRenderMode.Paged) {
            // CSS footnote extraction is a paged-media behavior. Continuous output
            // keeps the authored note in normal flow instead of dropping its body.
            style.FloatSide = "none";
            style.UnsupportedFloat = string.Empty;
        }
        RememberColumnSourceStyle(element, style);
        if (style.Display == "none") return;
        if (tag == "table") flattenInternalTableBoxes = style.Display == "inline";
        bool isBlock = HtmlRenderStyleResolver.IsBlockElement(element, style)
            && !(flattenInternalTableBoxes && IsInternalTableDisplay(style.Display));
        if (tag == "br") {
            runs.Add(new HtmlInlineRun("\u2028", inheritedStyle, inheritedLink, HtmlRenderStyleResolver.DescribeSource(element), inheritedPaintOffsetX, inheritedPaintOffsetY, element));
            return;
        }
        bool isPageFloat = TryGetPageFloatSide(element, style, out _);
        bool isColumnFloat = TryGetColumnEdgeFloatSide(element, style, out _, out _);
        // These boxes run LayoutElement, which owns their destination at the
        // laid-out box. An additional zero-flow marker would duplicate it.
        bool ownsAtomicDestination = !IsReplacedImageElementTag(tag) && tag is not ("math" or "ruby")
            && (!IsFormControlElement(tag) || UsesButtonChildLayout(element))
            && style.Display is "inline-block" or "inline-flex" or "inline-grid" or "inline-table";
        if (!isBlock && !isPageFloat && !isColumnFloat && !ownsAtomicDestination && IsFirstColumnSourceFragment(element)) {
            AddInlineNamedDestinationRun(element, style, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
        }
        if (!isPageFloat && !isColumnFloat) ReportUnsupportedFloatValues(element, style);
        ReportUnsupportedOverflowValues(element, style);
        ReportUnsupportedMultiColumnValues(element, style);
        RegisterInlineSemanticControls(element, style);
        if (style.Display == "inline" && TryGetColumnSourceFragment(element, out _)
            && !style.SemanticArtifact && !style.SemanticGroupRoleOverride.HasValue) {
            style = style.Clone();
            style.SemanticGroupRoleOverride = HtmlRenderSemanticGroupRole.Division;
        }
        string? link = inheritedLink;
        if (tag == "a") {
            link = ResolveSafeLink(element.GetAttribute("href"), element);
        }
        if (HtmlCssRunningElementParser.TryParsePosition(style.Position, out string runningElementName)) {
            AssignLogicalTextOrders(runs);
            runs.Add(new HtmlInlineRun(
                CaptureRunningElement(element, runningElementName, width, style, inheritedStyle, depth + 1),
                style,
                HtmlRenderStyleResolver.DescribeSource(element)));
            return;
        }
        if ((style.Position == "relative" || style.Position == "sticky") && style.ZIndex != "auto") {
            _inlineStackingElements.Add(element);
        }
        if (style.Position == "absolute" || style.Position == "fixed") {
            RegisterOutOfFlowElement(element.ParentElement ?? element, element, style, inheritedStyle, depth);
            runs.Add(new HtmlInlineRun(
                string.Empty,
                style,
                null,
                HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX,
                inheritedPaintOffsetY,
                element.ParentElement,
                element));
            return;
        }
        if (isColumnFloat) {
            AssignLogicalTextOrders(runs);
            if (TryAddColumnEdgeFloatRun(element, inheritedStyle, depth, style, link, runs)) return;
        }
        if (isPageFloat) {
            AssignLogicalTextOrders(runs);
            if (TryAddPageFloatRun(element, inheritedStyle, depth, style, link, runs)) return;
        }
        // Extracted figures own their destination at the moved box. A
        // rejected extraction retains the ordinary inline target in source flow.
        if ((isPageFloat || isColumnFloat) && !isBlock && !ownsAtomicDestination) {
            AddInlineNamedDestinationRun(element, style, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
        }
        if (style.FloatSide != "none") {
            if (style.FloatSide == "footnote") AddFootnoteRun(element, width, inheritedStyle, depth, style, runs);
            else {
                AssignLogicalTextOrders(runs);
                AddFloatingRun(element, width, inheritedStyle, depth, style, link, runs);
            }
            return;
        }

        if (IsInternalTableDisplay(style.Display) && !flattenInternalTableBoxes) ReportTableFormattingFallback(element, "orphan table boxes in an inline formatting context");

        if (isBlock) {
            AssignLogicalTextOrders(runs);
            HtmlRenderFlowBlock block = LayoutElement(element, width, style, inheritedStyle, depth + 1);
            runs.Add(new HtmlInlineRun(block, style, link, HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX, inheritedPaintOffsetY, element, isBlockInterruption: true));
            return;
        }

        if (IsFormControlElement(tag) && !UsesButtonChildLayout(element)) {
            AssignLogicalTextOrders(runs);
            if (!string.IsNullOrWhiteSpace(style.StringSet)) {
                runs.Add(new HtmlInlineRun(
                    element,
                    style,
                    HtmlRenderStyleResolver.DescribeSource(element)));
            }
            ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double controlOffsetX, out double controlOffsetY);
            double outerWidth = ResolveFormControlOuterWidth(element, style, width);
            HtmlRenderFlowBlock control = LayoutFormControl(element, outerWidth, style);
            control = ApplyElementPaintEffects(control, style, outerWidth, element, out _);
            var controlRun = new HtmlInlineRun(
                control,
                style,
                link,
                HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX + controlOffsetX,
                inheritedPaintOffsetY + controlOffsetY,
                element,
                isReplacedImage: true);
            int controlNodeId = GetSemanticNodeId(element);
            if (ShouldAssignNavigationNode(style)) {
                controlRun.AssignSemanticNode(style.SemanticRole, controlNodeId, ResolveBookmarkAnchorText(element, style));
            }
            AssignInlineSemanticGroup(controlRun, style, controlNodeId);
            runs.Add(controlRun);
            return;
        }

        if (tag == "ruby") {
            AssignLogicalTextOrders(runs);
            AddInlineRubyRun(
                element,
                width,
                inheritedStyle,
                depth,
                style,
                link,
                inheritedPaintOffsetX,
                inheritedPaintOffsetY,
                runs);
            return;
        }

        if (!IsReplacedImageElementTag(tag) && tag != "math" && style.Display == "inline-table") {
            AssignLogicalTextOrders(runs);
            HtmlRenderFlowBlock table = LayoutElement(element, width, style, inheritedStyle, depth + 1);
            runs.Add(new HtmlInlineRun(table, style, link, HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX, inheritedPaintOffsetY, element));
            return;
        }
        if (!IsReplacedImageElementTag(tag) && tag != "math" && style.Display == "inline-block") {
            AssignLogicalTextOrders(runs);
            AddInlineBlockRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (!IsReplacedImageElementTag(tag) && tag != "math" && style.Display == "inline-flex") {
            AssignLogicalTextOrders(runs);
            AddInlineFlexRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (!IsReplacedImageElementTag(tag) && tag != "math" && style.Display == "inline-grid") {
            AssignLogicalTextOrders(runs);
            AddInlineGridRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }

        if (style.UsedTransform != "none"
            || style.OpacityWasSpecified && (style.Opacity < 1D || style.UnsupportedOpacity.Length > 0)
            || style.ClipPath != "none"
            || HasInlineBoxPaint(style)) {
            _inlineStackingElements.Add(element);
        }

        if (!string.IsNullOrWhiteSpace(style.StringSet)) {
            runs.Add(new HtmlInlineRun(
                element,
                style,
                HtmlRenderStyleResolver.DescribeSource(element)));
        }

        ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double elementPaintOffsetX, out double elementPaintOffsetY);
        double paintOffsetX = inheritedPaintOffsetX + elementPaintOffsetX;
        double paintOffsetY = inheritedPaintOffsetY + elementPaintOffsetY;

        List<HtmlInlineRun>? semanticRuns = ShouldCollectSemanticInlineRuns(style)
            ? new List<HtmlInlineRun>()
            : null;
        IList<HtmlInlineRun> targetRuns = semanticRuns ?? runs;
        int firstInlineRun = targetRuns.Count;
        HtmlInlineEdgeScope? inlineEdgeScope = !IsReplacedImageElementTag(tag) && tag != "math"
            ? ResolveInlineEdgeScope(element, style, inheritedStyle, paintOffsetX) : null;
        HtmlInlineRun? inlineStrutMarker = null;
        if (inlineEdgeScope != null) targetRuns.Add(CreateInlineEdgeBoundary(inlineEdgeScope, true, link, paintOffsetX, paintOffsetY));
        else if (style.Display == "inline" && !IsReplacedImageElementTag(tag) && tag != "math") {
            inlineStrutMarker = new HtmlInlineRun(string.Empty, style, link, HtmlRenderStyleResolver.DescribeSource(element),
                paintOffsetX, paintOffsetY, element) { IsInlineStrutMarker = true };
            targetRuns.Add(inlineStrutMarker);
        }
        AddUnicodeBidiBoundaryRun(style, element, opening: true, paintOffsetX, paintOffsetY, targetRuns);
        if (IncludeColumnGeneratedContent(element, HtmlPseudoElementKind.Before))
            AddGeneratedInlineRun(element, HtmlPseudoElementKind.Before, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);

        if (IsReplacedImageElementTag(tag)) {
            if (IsInlineFrameElement(element)) {
                // The child document allocates logical text orders immediately.
                // Reserve preceding inline text before entering that document.
                AssignLogicalTextOrders(runs);
                if (!ReferenceEquals(targetRuns, runs)) AssignLogicalTextOrders(targetRuns);
            }
            AddInlineImageRun(element, style, link, paintOffsetX, paintOffsetY, targetRuns);
            AddUnicodeBidiBoundaryRun(style, element, opening: false, paintOffsetX, paintOffsetY, targetRuns);
            AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
            return;
        }
        if (tag == "math") {
            AssignLogicalTextOrders(runs);
            if (!ReferenceEquals(targetRuns, runs)) AssignLogicalTextOrders(targetRuns);
            if (TryAddInlineMathRun(element, width, style, link, paintOffsetX, paintOffsetY, targetRuns)) {
                AddUnicodeBidiBoundaryRun(style, element, opening: false, paintOffsetX, paintOffsetY, targetRuns);
                AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
                return;
            }
        }

        foreach (INode child in SelectedColumnChildNodes(element)) {
            CollectInlineRuns(child, width, containingHeight, style, link, depth + 1, paintOffsetX, paintOffsetY, targetRuns, flattenInternalTableBoxes);
        }

        if (IncludeColumnGeneratedContent(element, HtmlPseudoElementKind.After))
            AddGeneratedInlineRun(element, HtmlPseudoElementKind.After, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);
        AddUnicodeBidiBoundaryRun(style, element, opening: false, paintOffsetX, paintOffsetY, targetRuns);
        if (inlineEdgeScope != null) targetRuns.Add(CreateInlineEdgeBoundary(inlineEdgeScope, false, link, paintOffsetX, paintOffsetY));
        // Non-empty descendants carry their ancestor struts with the content
        // when wrapping. An extra opening marker would leave those struts on
        // the preceding line after all of the inline's text moved away.
        if (inlineStrutMarker != null && targetRuns.Skip(firstInlineRun + 1)
            .Any(run => !string.IsNullOrWhiteSpace(run.Text) || run.AtomicBlock != null || run.LeaderPattern != null))
            targetRuns.Remove(inlineStrutMarker);
        if (style.Display == "inline") CaptureInlineStrutStyles(targetRuns, firstInlineRun, style);
        AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
    }

    private static void AddUnicodeBidiBoundaryRun(
        HtmlRenderBoxStyle style,
        IElement owner,
        bool opening,
        double paintOffsetX,
        double paintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        string control = GetUnicodeBidiBoundary(style.UnicodeBidi, style.Direction, opening);
        if (control.Length == 0) return;
        runs.Add(new HtmlInlineRun(
            control,
            style,
            null,
            HtmlRenderStyleResolver.DescribeSource(owner),
            paintOffsetX,
            paintOffsetY,
            owner));
    }

    private static string GetUnicodeBidiBoundary(string unicodeBidi, string direction, bool opening) {
        bool rtl = string.Equals(direction, "rtl", StringComparison.Ordinal);
        switch (unicodeBidi) {
            case "embed": return opening ? rtl ? "\u202B" : "\u202A" : "\u202C";
            case "bidi-override": return opening ? rtl ? "\u202E" : "\u202D" : "\u202C";
            case "isolate": return opening ? rtl ? "\u2067" : "\u2066" : "\u2069";
            case "isolate-override":
                return opening
                    ? (rtl ? "\u2067\u202E" : "\u2066\u202D")
                    : "\u202C\u2069";
            case "plaintext": return opening ? "\u2068" : "\u2069";
            default: return string.Empty;
        }
    }

    private static bool ShouldAssignNavigationNode(HtmlRenderBoxStyle style) =>
        !style.SemanticArtifact
        && (HtmlRenderHeading.TryGetLevel(style.SemanticRole, out _)
            || style.BookmarkLevelSpecified);

    private static bool ShouldCollectSemanticInlineRuns(HtmlRenderBoxStyle style) =>
        ShouldAssignNavigationNode(style)
        || style.SemanticArtifact
        || style.SemanticGroupRoleOverride.HasValue;

    private void AppendSemanticInlineRuns(
        IElement element,
        HtmlRenderBoxStyle style,
        IReadOnlyList<HtmlInlineRun>? semanticRuns,
        ICollection<HtmlInlineRun> destination,
        string? link,
        double paintOffsetX,
        double paintOffsetY) {
        if (semanticRuns == null) return;
        int nodeId = GetSemanticNodeId(element);
        string bookmarkAnchorText = ResolveBookmarkAnchorText(element, style);
        if (ShouldAssignNavigationNode(style)
            && IsFirstColumnSourceFragment(element)
            && !style.BookmarkSuppressed
            && bookmarkAnchorText.Length > 0) {
            string source = HtmlRenderStyleResolver.DescribeSource(element);
            var anchor = new HtmlRenderBookmarkAnchor(nodeId, bookmarkAnchorText, 0D, 0D, 0.01D, 0.01D, 0, source);
            var markerBlock = new HtmlRenderFlowBlock(
                0D,
                0.01D,
                new[] { anchor },
                HtmlPageBreakTarget.None,
                HtmlPageBreakTarget.None,
                false,
                source);
            var markerRun = new HtmlInlineRun(
                markerBlock,
                style,
                link,
                source,
                paintOffsetX,
                paintOffsetY,
                element,
                isFlowMarker: true);
            markerRun.AssignSemanticNode(style.SemanticRole, nodeId, bookmarkAnchorText);
            if (semanticRuns.Count == 0) AssignInlineSemanticGroup(markerRun, style, nodeId);
            destination.Add(markerRun);
            if (semanticRuns.Count == 0) return;
        }
        foreach (HtmlInlineRun run in semanticRuns) {
            if (ShouldAssignNavigationNode(style)) run.AssignSemanticNode(style.SemanticRole, nodeId, bookmarkAnchorText);
            AssignInlineSemanticGroup(run, style, nodeId);
            destination.Add(run);
        }
    }

    private static void AssignInlineSemanticGroup(HtmlInlineRun run, HtmlRenderBoxStyle style, int semanticNodeId) {
        string structureElementKey = "html-inline:" + semanticNodeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (style.SemanticArtifact) run.AssignInlineSemanticGroup(HtmlRenderSemanticGroupRole.Artifact, structureElementKey);
        else if (style.SemanticGroupRoleOverride.HasValue) run.AssignInlineSemanticGroup(style.SemanticGroupRoleOverride.Value, structureElementKey);
    }

}
