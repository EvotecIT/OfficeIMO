namespace OfficeIMO.Pdf;

internal static class PdfFlowNestingRules {
    /// <summary>Checks static column content recursively without changing its layout or semantics.</summary>
    public static bool IsColumnFlowSupported(IPdfBlock block) => block switch {
        ContainerBlock container => container.Blocks.All(IsColumnFlowSupported),
        SemanticBlock semantic => semantic.Blocks.All(IsColumnFlowSupported),
        FlowBlock flow => !flow.IsReplayable && flow.StaticBlocks != null &&
                          flow.Options.ShowIf == null && flow.Options.MinimumRemainingHeight == 0D &&
                          flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Continue &&
                          flow.StaticBlocks.All(IsColumnFlowSupported),
        _ => IsColumnFlowPrimitive(block)
    };

    public static bool IsColumnFlowPrimitive(IPdfBlock block) =>
        block is HeadingBlock or RichParagraphBlock or BulletListBlock or NumberedListBlock or
        TableBlock or HorizontalRuleBlock or ImageBlock or ShapeBlock or DrawingBlock or
        TextFieldBlock or CheckBoxBlock or ChoiceFieldBlock or RadioButtonGroupBlock or
        TextAnnotationBlock or FreeTextAnnotationBlock or HighlightAnnotationBlock or BookmarkBlock or SpacerBlock;
}
