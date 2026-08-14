using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock ApplyElementSemantics(HtmlRenderFlowBlock block, IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
        if (ShouldAssignNavigationNode(style) && !style.BookmarkSuppressed) {
            int nodeId = GetSemanticNodeId(element);
            string anchorText = ResolveBookmarkAnchorText(element, style);
            if (anchorText.Length > 0 && !ContainsNavigationFragment(block.Visuals, nodeId)) {
                block = block.WithVisuals(block.Visuals.Concat(new HtmlRenderVisual[] {
                    new HtmlRenderBookmarkAnchor(
                        nodeId,
                        anchorText,
                        style.MarginLeft,
                        style.MarginTop,
                        Math.Max(0.01D, block.Width - style.MarginLeft - style.MarginRight),
                        Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom),
                        block.Visuals.Count,
                        HtmlRenderStyleResolver.DescribeSource(element))
                }));
            }
        }
        HtmlRenderFlowBlock listBlock = style.SemanticArtifact || style.SemanticGroupRoleOverride.HasValue
            ? block
            : ApplyListSemantics(block, element);
        HtmlRenderSemanticGroupRole role;
        if (style.SemanticArtifact) role = HtmlRenderSemanticGroupRole.Artifact;
        else if (style.SemanticGroupRoleOverride.HasValue) role = style.SemanticGroupRoleOverride.Value;
        else if (!TryResolveSemanticGroupRole(element.TagName, out role)) return listBlock;
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

    private HtmlRenderFlowBlock ApplySpecializedElementSemantics(HtmlRenderFlowBlock block, IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
        if (ShouldAssignNavigationNode(style) && !style.BookmarkSuppressed) {
            int nodeId = GetSemanticNodeId(element);
            string anchorText = ResolveBookmarkAnchorText(element, style);
            if (anchorText.Length > 0) {
                block = block.WithVisuals(block.Visuals.Concat(new HtmlRenderVisual[] {
                    new HtmlRenderBookmarkAnchor(
                        nodeId,
                        anchorText,
                        0D,
                        0D,
                        Math.Max(0.01D, block.Width),
                        Math.Max(0.01D, block.Height),
                        block.Visuals.Count,
                        HtmlRenderStyleResolver.DescribeSource(element))
                }));
            }
        }
        if (!style.SemanticArtifact && !style.SemanticGroupRoleOverride.HasValue) return block;
        HtmlRenderSemanticGroupRole role = style.SemanticArtifact
            ? HtmlRenderSemanticGroupRole.Artifact
            : style.SemanticGroupRoleOverride!.Value;
        return block.WithVisuals(new[] {
            new HtmlRenderSemanticGroup(
                role,
                0D,
                0D,
                Math.Max(0.01D, block.Width),
                Math.Max(0.01D, block.Height),
                block.Visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        });
    }

    private static string ResolveBookmarkAnchorText(IElement element, HtmlRenderBoxStyle style) {
        if (!string.IsNullOrWhiteSpace(style.BookmarkLabel)) return style.BookmarkLabel!.Trim();
        string text = element.TextContent?.Trim() ?? string.Empty;
        if (text.Length > 0) return text;
        return (element.GetAttribute("aria-label")
            ?? element.GetAttribute("alt")
            ?? element.GetAttribute("title")
            ?? string.Empty).Trim();
    }

    private HtmlRenderVisual ApplyInlineElementSemantics(HtmlRenderVisual visual, HtmlInlineRun run) {
        if (!run.InlineSemanticGroupRole.HasValue) return visual;
        HtmlRenderSemanticGroupRole role = run.InlineSemanticGroupRole.Value;
        if (visual is HtmlRenderSemanticGroup existing && existing.Role == role) return visual;
        return new HtmlRenderSemanticGroup(
            role,
            visual.X,
            visual.Y,
            Math.Max(0.01D, visual.Width),
            Math.Max(0.01D, visual.Height),
            new[] { visual },
            visual.PaintOrder,
            run.Source,
            layoutY: visual.LayoutY,
            structureElementKey: run.InlineSemanticGroupKey);
    }

    private static bool ContainsBookmarkAnchor(IEnumerable<HtmlRenderVisual> visuals, int semanticNodeId) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderBookmarkAnchor anchor && anchor.SemanticNodeId == semanticNodeId) return true;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals
                : null;
            if (children != null && ContainsBookmarkAnchor(children, semanticNodeId)) return true;
        }
        return false;
    }

    private static bool ContainsNavigationFragment(IEnumerable<HtmlRenderVisual> visuals, int semanticNodeId) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderBookmarkAnchor anchor && anchor.SemanticNodeId == semanticNodeId) return true;
            if (visual is HtmlRenderText text && text.SemanticNodeId == semanticNodeId) return true;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals
                : null;
            if (children != null && ContainsNavigationFragment(children, semanticNodeId)) return true;
        }
        return false;
    }

    private void RegisterInlineSemanticControls(IElement element, HtmlRenderBoxStyle style) {
        RegisterBookmark(element, style);
        ReportUnsupportedSemanticTag(element, style);
    }

    private void ReportUnsupportedSemanticTag(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedSemanticTag.Length == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.PdfSemanticTagUnsupported,
            "A PDF semantic tag override was invalid and automatic HTML semantics were used.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            "-officeimo-pdf-tag-type=" + style.UnsupportedSemanticTag,
            OfficeConversionLossKind.Approximation);
    }

    private void RegisterBookmark(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedBookmark.Length > 0) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.BookmarkValueUnsupported,
                "A CSS bookmark value used automatic heading navigation.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(element),
                style.UnsupportedBookmark,
                OfficeConversionLossKind.Approximation);
        }
        bool automaticHeading = HtmlRenderHeading.TryGetLevel(style.SemanticRole, out int headingLevel);
        if (!automaticHeading && !style.BookmarkLevelSpecified && style.BookmarkLabel == null && style.BookmarkState == HtmlRenderBookmarkState.Default) return;
        int nodeId = GetSemanticNodeId(element);
        int level = style.BookmarkLevel ?? (automaticHeading ? headingLevel : 1);
        _bookmarkDefinitions[nodeId] = new HtmlRenderBookmarkDefinition(level, style.BookmarkLabel, style.BookmarkState, style.BookmarkSuppressed, GetDocumentOrder(element));
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
