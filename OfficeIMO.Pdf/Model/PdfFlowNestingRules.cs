namespace OfficeIMO.Pdf;

internal static class PdfFlowNestingRules {
    /// <summary>Checks static column content recursively without changing its layout or semantics.</summary>
    public static bool IsColumnFlowSupported(IPdfBlock block, bool allowTableOfContents = false) => block switch {
        TableOfContentsBlock => allowTableOfContents,
        ContainerBlock container => container.Blocks.All(child => IsColumnFlowSupported(child, allowTableOfContents)),
        SemanticBlock semantic => semantic.Blocks.All(child => IsColumnFlowSupported(child, allowTableOfContents)),
        FlowBlock flow => !flow.IsReplayable && flow.StaticBlocks != null &&
                          flow.Options.ShowIf == null && flow.Options.MinimumRemainingHeight.ToPoints() == 0D &&
                          flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Continue &&
                          flow.StaticBlocks.All(child => IsColumnFlowSupported(child, allowTableOfContents)),
        _ => IsColumnFlowPrimitive(block)
    };

    public static bool IsColumnFlowPrimitive(IPdfBlock block) =>
        block is HeadingBlock or RichParagraphBlock or PdfListBlock or
        TableBlock or HorizontalRuleBlock or ImageBlock or ShapeBlock or DrawingBlock or
        TextFieldBlock or CheckBoxBlock or ChoiceFieldBlock or RadioButtonGroupBlock or
        TextAnnotationBlock or FreeTextAnnotationBlock or HighlightAnnotationBlock or BookmarkBlock or SpacerBlock;
}
